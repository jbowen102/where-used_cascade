import os
import argparse     # Used to parse optional command-line arguments

import class_def
from platforms import platform_dict

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory


parser = argparse.ArgumentParser(description="Program to automate recursive "
                                                        "where-used analyses")
parser.add_argument("-t", "--target-parts", help="Specify which parts are "
                        "of primary interest - comma-separated (no spaces).",
                                                    type=str, default=False)
parser.add_argument("-v", "--verbose", help="Include additional output for "
                                        "diagnosis.", action="store_true")
# https://www.programcreek.com/python/example/748/argparse.ArgumentParser
args = parser.parse_args()


AllParts = class_def.PartGroup(target_part_str=args.target_parts)
AllParts.import_platforms(platform_dict)

# AllParts.import_all_reports(report_type="SAPTC")
# AllParts.import_all_reports(report_type="SAP_multi_w")
# AllParts.import_all_reports(report_type="SAP_multi_BOM")

# AllParts.get_target_obs_status()
# AllParts.print_obs_status_trace()

# TreeViz = class_def.TreeGraph(AllParts, target_group_only=True, printout=True)
# TreeViz = class_def.TreeGraph(AllParts, target_group_only=True)
# TreeViz = class_def.TreeGraph(AllParts, target_group_only=True, include_desc=False)
# TreeViz = class_def.TreeGraph(AllParts)
# TreeViz.export_graph()

# AllParts.export_parts_set(omit_platforms=True)

# Union BOM
# AllParts.import_all_reports(report_type="SAP_multi_BOM", bom_union=True)
AllParts.import_all_reports(report_type="SAP_multi_BOM", bom_union=True,
                                             import_subdir="previous_multi-BOMs")

# hack to eliminate already-seen parts
# AllParts2 = class_def.PartGroup(target_part_str=args.target_parts)
# AllParts2.import_platforms(platform_dict)
# AllParts2.import_all_reports(report_type="SAP_multi_BOM",
#                                             import_subdir="previous_multi-BOMs")
# for Part_i in AllParts.get_parts():
#     if AllParts2.get_part(Part_i.get_pn()):
#         AllParts.union_bom.discard(Part_i)

AllParts.export_parts_set(pn_set=AllParts.get_union_bom(), omit_platforms=True)
