import os

import class_def
from platforms import platform_dict

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory


AllParts = class_def.PartGroup()
AllParts.import_platforms(platform_dict)

AllParts.import_report("")
AllParts.import_report("")

# for PartNum in AllParts.get_parts():
#     # print("\n%s" % PartNum)
#     print("%s has parents %r" % (PartNum, PartNum.get_parents()))
#     print("\tCan OBS? %r\n" % (PartNum.get_obs_status()))
#     # print("%s\n\tCan OBS? %r\n" % (PartNum, PartNum.get_obs_status()))
#     # print("\t%r\n" % PartNum.get_obs_status())
#     # print("%r" % PartNum)

AllParts.print_all_obs_status()
