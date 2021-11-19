import os
import csv

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory

class Part(object):
    """Object to represent a part, assy, or mod.
    """
    def __init__(self, part_num, name):
        self.part_num = part_num
        self.name = name
        self.Parents = set()

    def get_obs_status(self):
        # If the name of the part is OBS, don't bother w/ parent query.
        if len(self.name) > 3 and "OBS" in self.name[:4].upper():
            return True

        self.can_obs = True
        for Parent_i in self.Parents:
            print("%s: looking for status of parent %s" % (self.part_num, Parent_i))
            parent_status = Parent_i.get_obs_status()
            print("\t%s has parent %s - can obs? %r" % (self.part_num, Parent_i,
                                                                parent_status))
            if parent_status == False:
                self.can_obs = False
                break
        # If no parents in set, results in self.can_obs = True as it should.
        return self.can_obs

    def add_parent(self, Parent_i):
        self.Parents.add(Parent_i)

    def get_parent(self, parent_num):
        for Parent_i in self.Parents:
            if parent_num == Parent_i.get_pn():
                return Parent_i
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
        print("Importing platforms...")
        for platform in platform_dict:
            platform_pn = platform.split("-")[0]
            platform_desc = platform[len(platform_pn)+1:]
            platform_obs = platform_dict[platform]
            print("\tAdding platform %s to group" % platform_pn)
            self.add_part(Platform(platform_pn, platform_desc, platform_obs))
        print("...done")

    def add_part(self, Part_i):
        self.Parts.add(Part_i)

    def get_part(self, part_num):
        for Part_i in self.Parts:
            if part_num == Part_i.get_pn():
                return Part_i
        return False # only happens if no match found in loop.

    def get_parts(self):
        return self.Parts

    def print_all_obs_status(self):
        print("")
        for PartNum in self.get_parts():
            if PartNum.__class__.__name__ == "Platform":
                print("%s is a platform" % PartNum)
            else:
                print("%s has parents %r" % (PartNum, PartNum.get_parents()))
            print("\tCan OBS? %r\n" % (PartNum.get_obs_status()))

    def __repr__(self):
        return "PartsGroup object: %s" % str(self.Parts)

    def import_report(self, file_name):
        """Read in where-used result in import directory.
        """
        import_path = os.path.join(self.import_dir, file_name)

        with open(import_path, "r") as import_file:
            print("\nReading data from %s" % import_path)
            import_file_it = csv.reader(import_file)

            raw_import_dict = {}

            for i, import_row in enumerate(import_file_it):
                if i == 1:
                    part_num = import_row[2]

                    assert import_row[0] == "Material:", ("Expected "
                                    "'Material:' in first column of row %s. "
                                    "Check formatting in %s." % (i, file_name))
                    assert len(part_num) >= 6, ("Found less than 6 digits "
                       "where part number should be in third column of row %s. "
                                    "Check formatting in %s." % (i, file_name))

                if i == 2:
                    part_desc = import_row[2]

                    assert import_row[0] == "Description:", ("Expected "
                                    "'Description:' in first column of row %s. "
                                    "Check formatting in %s." % (i, file_name))
                    assert len(part_desc) > 0, ("Found empty cell where "
                      "description string should be in third column of row %s. "
                                    "Check formatting in %s." % (i, file_name))

                    if self.get_part(part_num) == False:
                        ThisPart = Part(part_num, part_desc)
                        print("\tAdding main part %s to group" % ThisPart)
                        self.add_part(ThisPart)
                    else:
                        print("\tMain part %s already in group" % part_num)
                        ThisPart = self.get_part(part_num)
                    print("\t\tMain part %s has parent set %r" % (ThisPart,
                                                        ThisPart.get_parents()))
                    print("")
                    # print("\t\tGroup: %r\n" % self.get_parts())

                if i == 6:
                    assert import_row[3] == "Component", ("Expected "
                                    "'Component' in first column of row %s. "
                                    "Check formatting in %s." % (i, file_name))
                    assert import_row[4] == "Component Description", (
                        "Expected 'Component Description' in first column of "
                            "row %s. Check formatting in %s." % (i, file_name))

                if i > 6 and "End of Report" not in import_row[0]:
                    parent_num = import_row[3]
                    parent_desc = import_row[4]

                    assert len(parent_num) >= 6, ("Found less than 6 digits "
                      "where part number should be in fourth column of row %s. "
                                    "Check formatting in %s." % (i, file_name))
                    assert len(parent_desc) > 0, ("Expected "
                        "'Component Description' in first column of row %s. "
                                    "Check formatting in %s." % (i, file_name))

                    # Create and add this part to the group if not already in
                    # the Parts set.
                    if self.get_part(parent_num) == False:
                        NewParent = Part(parent_num, parent_desc)
                        print("\tAdding part %s to group" % NewParent)
                        self.add_part(NewParent)
                    else:
                        print("\tPart %s already in group" % parent_num)
                        NewParent = self.get_part(parent_num)
                    print("\t\tPart %s has parent set %r" % (NewParent,
                                                       NewParent.get_parents()))
                    # print("\t\tGroup: %r" % self.get_parts())

                    # Add this part as a parent if not already in
                    # the Parents set.
                    if ThisPart.get_parent(parent_num) == False:
                        print("\tAdding %s as parent of part %s" % (NewParent,
                                                                      ThisPart))
                        ThisPart.add_parent(NewParent)
                        # print("\t\tPart %s has parent set %r" % (ThisPart,
                        #                                 ThisPart.get_parents()))

                    for Part_i in self.get_parts():
                        if Part_i.__class__.__name__ != "Platform":
                            print("\t\tPart %s has parent set %r" % (Part_i,
                                                          Part_i.get_parents()))
                    print("")

        print("\tGroup: %r" % self.get_parts())
        print("...done")


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
