import os

import class_def
from platforms import platform_dict

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory


AllParts = set()

for platform in platform_dict:
    AllParts.add(class_def.Platform(platform, platform_dict[platform]))

for Part in AllParts:
    print("%r\n\tCan OBS? %r\n" % (Part, Part.get_obs_status()))
    # print("%s: %r" % (Part, Part.get_obs_status()))
# print(AllParts)
