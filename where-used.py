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
# https://www.programcreek.com/python/example/748/argparse.ArgumentParser
args = parser.parse_args()

assert args.mode, "Need to pass mode argument."

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
    """Reads in SAP multi-level where-used report(s), ignores target parts.
    Exports graph showing structure of BOM along with can-obsolete coloring.
    """
    AllParts.import_all_reports(report_type="SAP_multi_w")
    # AllParts.get_target_obs_status()
    # AllParts.print_obs_status_trace()
    TreeViz = class_def.TreeGraph(AllParts, target_group_only=True,
                            printout=args.printout, exclude_desc=args.compact)
    TreeViz.export_graph()

elif args.mode.lower() == "union":
    """Reads in SAP multi-level BOM(s), reads in target parts.
    Exports list of target parts and every part used in any level below the target parts.
    """
    AllParts.import_all_reports(report_type="SAP_multi_BOM")

    # hack to eliminate already-seen parts for successive large runs
    # AllParts2 = class_def.PartGroup(target_part_str=args.target_parts)
    # AllParts2.import_platforms(platform_dict)
    # AllParts2.import_all_reports(report_type="SAP_multi_BOM",
    #                                             import_subdir="previous_multi-BOMs")
    # for Part_i in AllParts.get_parts():
    #     if AllParts2.get_part(Part_i.get_pn()):
    #         AllParts.union_bom.discard(Part_i)

    AllParts.export_parts_set(pn_set=AllParts.get_union_bom(), omit_platforms=True)
    # Export union bom w/ platform applications:
    # AllParts.export_parts_set(pn_set=AllParts.get_union_bom(),
    #                                     omit_platforms=True, platform_app=True)

elif args.mode.lower() == "platform":
    """Reads in SAP platform multi-level BOM(s), reads in target parts.
    Exports list of target parts along with which platforms each is used in.
    """
    AllParts.import_all_reports(report_type="SAP_multi_BOM")
    AllParts.import_target_parts()

    AllParts.export_parts_set(pn_set=AllParts.get_target_parts(),
                                        omit_platforms=True, platform_app=True)
    # Export union bom w/ platform applications:
    # AllParts.export_parts_set(pn_set=AllParts.get_union_bom(),
    #                                     omit_platforms=True, platform_app=True)

elif args.mode.lower() == "union_loop":
    """Reads in SAP multi-level BOM(s).
    Repeatedly prompts user for individual target part to create union BOM for.
    Exports list of target part and every part used in any level below the target part.
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
