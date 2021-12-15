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

AllParts.import_all_reports()
# Test if any needed report is missing.
AllParts.find_missing_reports()


AllParts.get_target_obs_status()
# AllParts.print_obs_status_trace()

AllParts.export_tree_viz(target_group_only=True)
# AllParts.export_tree_viz()
