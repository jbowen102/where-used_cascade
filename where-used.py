import os
import argparse     # Used to parse optional command-line arguments

import class_def
from platforms import platform_dict

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory


parser = argparse.ArgumentParser(description="Program to automate recursive "
                                                        "where-used analyses")
parser.add_argument("-v", "--verbose", help="Include additional output for "
                                            "diagnosis.", action="store_true")
parser.add_argument("-m", "--mode", help="Specify which mode to run program in "
                    "('single', 'multi', 'union', 'platform'). 'union' and "
        "'platform' types use SAP multi-level BOM(s)", type=str, default=None)
parser.add_argument("-gp", "--printout", help="Specify that graph should be "
                "created in printout mode (sparse color). "
                "Only valid in 'single' or 'multi' modes.", action="store_true")
parser.add_argument("-gc", "--compact", help="Specify that graph should be "
                "created with compact nodes (part descriptions omitted). "
                "Only valid in 'single' or 'multi' modes.", action="store_true")
parser.add_argument("-t", "--target-part", help="Pass in single target part "
                        "to use in place of target_parts.txt contents.",
                                                        type=str, default=None)
parser.add_argument("-e", "--exclude-obs", help="Exclude obsolete and orphaned "
                "parts from multi-level where-used graph.", action="store_true")
# https://www.programcreek.com/python/example/748/argparse.ArgumentParser
args = parser.parse_args()

assert args.mode, "Need to pass mode argument."
assert args.mode in ["single", "multi", "union", "platform", "platform_union",
                                        "assy_list", "union_loop", "bom_vis"]
if args.exclude_obs:
    assert args.mode == "multi", "-e flag can only be used with multi mode."
if args.target_part:
    with open(class_def.TARGET_PARTS_PATH, "r") as target_parts_file:
        # Display contents about to be overwritten.
        print("Previous %s contents:" % os.path.basename(class_def.TARGET_PARTS_PATH))
        print(target_parts_file.read())

    with open(class_def.TARGET_PARTS_PATH, "w") as target_parts_file:
        target_parts_file.write(args.target_part.upper())

AllParts = class_def.PartGroup()
AllParts.import_platforms(platform_dict)

if args.mode.lower() == "single":
    """Reads in TC-SAP single-level where-used report(s), builds structure from
    target parts upward, prompting user for more single-level reports (or orphan
    determination) as needed.
    Exports graph showing structure of BOM along with can-obsolete coloring.
    """
    AllParts.import_all_reports(report_type="SAPTC")
    # AllParts.get_target_obs_status()
    # AllParts.print_obs_status_trace()
    TreeViz = class_def.TreeGraph(AllParts, target_group_only=True,
                            printout=args.printout, exclude_desc=args.compact)
    TreeViz.export_graph()

elif args.mode.lower() == "multi":
    """Reads in SAP multi-level where-used report(s), reads in target parts.
    Exports graph showing where-used hierarchy along with can-obsolete coloring.
    """
    AllParts.import_all_reports(report_type="SAP_multi_w")

    TreeViz = class_def.TreeGraph(AllParts, target_group_only=True,
                            printout=args.printout, exclude_desc=args.compact,
                                                exclude_obs=args.exclude_obs)
    TreeViz.export_graph()

elif args.mode.lower() == "union":
    """Reads in SAP multi-level BOM(s), reads in target parts.
    Exports list of target parts and every part used in any level below the
    target parts.
    Program will report any target parts not found in multi-BOMs (and thus not
    expanded).
    """
    AllParts.import_all_reports(report_type="SAP_multi_BOM")

    #### TEMP - used to see if mods/parts being used on new platforms include
    ####        any parts I'm obsoleting
    # NewParts = AllParts.get_union_bom() # new stuff in target_parts.txt
    # AllParts.target_Parts = set()
    # input("\n\nReplace target parts") # put parts planning to obs in target_parts.txt
    # AllParts.import_target_parts(parts_update=False)
    # ObsParts = AllParts.get_target_parts()
    # print(NewParts.intersection(ObsParts))
    #### TEMP

    AllParts.export_parts_set(pn_set=AllParts.get_union_bom(), omit_platforms=True)

elif args.mode.lower() == "platform":
    """Reads in SAP multi-level BOM(s), reads in target parts.
    Exports list of target parts along with which platforms each is used in.
    Set of SAP multi-level BOMs should include every platform user wants to see
    in results, e.g. need platform 666111 multi-BOM if you want that platform
    to show up next to parts it uses.
    Non-platform multi-BOMs in import folder are ignored.
    """
    AllParts.import_all_reports(report_type="SAP_multi_BOM")
    AllParts.import_target_parts()

    AllParts.export_parts_set(pn_set=AllParts.get_target_parts(),
                                        omit_platforms=True, platform_app=True)

elif args.mode.lower() == "platform_union":
    """Reads in SAP multi-level BOM(s), reads in target parts.
    Exports list of target parts and every part used in any level below the
    target parts, along with which platforms each is used in.
    Set of SAP multi-level BOMs should include every platform user wants to see
    in results, e.g. need platform 666111 multi-BOM if you want that platform
    to show up next to parts it uses.
    Non-platform multi-BOMs in import folder are used only if they contain a
    target part. In that case, the target part's BOM is read from the multi-BOM
    so its constituents can be unioned and included in export (along w/ platform
    applications).
    Program will report any target parts not found in multi-BOMs (and thus not
    expanded).
    """
    AllParts.import_all_reports(report_type="SAP_multi_BOM")

    # Export union bom w/ platform applications:
    AllParts.export_parts_set(pn_set=AllParts.get_union_bom(),
                                        omit_platforms=True, platform_app=True)

elif args.mode.lower() == "assy_list":
    """Reads in SAP multi-level where-used report(s), reads in target parts.
    Returns list of P/Ns in where-used hierarchy that aren't platforms or mods.
    """
    AllParts.import_all_reports(report_type="SAP_multi_w")

    print("")
    for TargetPart in AllParts.get_target_parts():
        assy_set = TargetPart.get_parents_above(assy_only=True)
        # assy_set = TargetPart.get_parents_above()
        print("%s: " % TargetPart)
        for Part_i in assy_set:
            print("\t%s - %s" % (Part_i, Part_i.get_name()))


elif args.mode.lower() == "union_loop":
    """Reads in SAP multi-level BOM(s).
    Repeatedly prompts user for individual target part to create union BOM for.
    Exports list of target part and every part used in any level below the
    target part.
    """
    AllParts.import_all_reports(report_type="SAP_multi_BOM")

    while True:
        print("\nEnter P/N")
        pn = input("> ")
        if not AllParts.get_part(pn):
            print("P/N not found in set.")
            continue
        else:
            # Manually edit target_Parts so get_union_bom() ignores txt file.
            AllParts.target_Parts = set({AllParts.get_part(pn)})
            AllParts.export_parts_set(pn_set=AllParts.get_union_bom(), omit_platforms=True)

elif args.mode.lower() == "bom_vis":
    """Reads in SAP multi-level BOM(s), reads in target parts.
    Exports graph showing BOM hierarchy along with can-obsolete coloring.
    Can't have any multi-level BOMs in the import folder that you don't want on
    the graph.
    """
    AllParts.import_all_reports(report_type="SAP_multi_BOM")
    # AllParts.import_target_parts()

    TreeViz = class_def.TreeGraph(AllParts, target_group_only=False,
                            printout=args.printout, exclude_desc=args.compact)
    TreeViz.export_graph()
