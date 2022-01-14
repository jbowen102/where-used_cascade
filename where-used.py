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
                    "('single', 'multi', 'union'). 'union' type uses SAP "
                                "multi-level BOM(s)", type=str, default=None)
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
    AllParts.import_all_reports(report_type="SAPTC")
    # AllParts.get_target_obs_status()
    # AllParts.print_obs_status_trace()
    TreeViz = class_def.TreeGraph(AllParts, target_group_only=True,
                            printout=args.printout, exclude_desc=args.compact)
    TreeViz.export_graph()
elif args.mode.lower() == "multi":
    AllParts.import_all_reports(report_type="SAP_multi_w")
    # AllParts.get_target_obs_status()
    # AllParts.print_obs_status_trace()
    TreeViz = class_def.TreeGraph(AllParts, target_group_only=True,
                            printout=args.printout, exclude_desc=args.compact)
    TreeViz.export_graph()
elif args.mode.lower() == "union":
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
