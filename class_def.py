import os
import csv

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory

class Part(object):
    """Object to represent a part, assy, or mod.
    """
    def __init__(self, part_num, name, Parents=set()):
        self.part_num = part_num
        self.name = name
        self.Parents = Parents

    def get_obs_status(self):
        # If the name of the part is OBS, don't bother w/ parent query.
        if len(self.name) > 3 and "OBS" in self.name[:4].upper():
            return True

        self.can_obs = True
        for Parent in self.Parents:
            parent_status = Parent.get_obs_status()
            if parent_status == False:
                self.can_obs = False
                break
        # If no parents in set, results in self.can_obs = True as it should.
        return self.can_obs

    def add_parent(self, Parent):
        self.Parents.add(Parent)

    def get_parent(self, parent_num):
        for Parent in self.Parents:
            if parent_num == Parent.get_pn():
                return Parent
        return False # only happens if no match found in loop.

    def get_parents(self):
        return self.Parents

    def get_pn(self):
        return self.part_num

    def get_name(self):
        return self.name

    def __str__(self):
        return self.part_num

    def __repr__(self):
        return "Part object: %s" % self.part_num


class Platform(Part):
    """Object to represent a platform. Inherits from Part class.
    """
    def __init__(self, part_num, name, can_obs):
        self.part_num = part_num
        self.name = name
        self.can_obs = can_obs
        self.Parents = None

    def get_obs_status(self):
        return self.can_obs

    def __repr__(self):
        return "Platform object: %s" % self.part_num


class PartGroup(object):
    def __init__(self, starting_set=set()):
        self.import_dir = os.path.join(SCRIPT_DIR, "import")
        self.Parts = starting_set

    def import_platforms(self, platform_dict):
        """Read in platform data from given dictionary (key is PN and value is
        True/False for can_obs).
        """
        for platform in platform_dict:
            self.Parts.add(Platform(platform, platform_dict[platform]))

    def get_part(self, part_num):
        for Part in self.Parts:
            if part_num == Part.get_pn():
                return Part
        return False # only happens if no match found in loop.

# testing
# Pltfm1 = Platform("658237", True)
# # print(Pltfm1, ": ", Pltfm1.get_obs_status())
# Pltfm2 = Platform("658244", False)
# Pltfm3 = Platform("10016534", True)
#
# Part3 = Part("10013547", Parents=set({Pltfm1, Pltfm3}))
# print(Part3, ": ", Part3.get_obs_status())
#
# Part4 = Part("654981", Parents=set({Pltfm2}))
# print(Part4, ": ", Part4.get_obs_status())
#
# Part3.add_parent(Part4)
# print(Part3, ": ", Part3.get_obs_status())
