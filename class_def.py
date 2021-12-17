print("Importing modules...")
import os
import csv
from datetime import datetime

import pandas as pd
import pydot
print("...done\n")

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory

class Part(object):
    """Object to represent a part, assy, or mod.
    """
    def __init__(self, part_num, name=""):
        self.part_num = part_num
        self.name = name
        self.Parents = set()

        # Establish if part has "OBS" prefix in SAP
        if self.name and len(self.name) > 3 and "OBS" in self.name[:4].upper():
            self.obs_disp = True
        else:
            self.obs_disp = False

        # Initialize variable indicating where-used results are present (until
        # determined otherwise)
        self.orphan = False
        self.report_name = None

    def set_report_name(self, report_name):
        self.report_name = report_name

    def get_report_name(self):
        return self.report_name

    def get_obs_disp(self):
        return self.obs_disp

    def get_obs_status(self, silent=False):
        # If the name of the part is OBS, don't bother w/ parent query.
        if self.obs_disp:
            return True

        self.can_obs = True
        for Parent_i in self.Parents:
            if not silent:
                print("%s: looking for status of parent %s" % (self.part_num,
                                                                    Parent_i))
            parent_status = Parent_i.get_obs_status(silent)
            if not silent:
                print("\t%s has parent %s - can obs? %r" % (self.part_num,
                                                    Parent_i, parent_status))
            if parent_status == False:
                self.can_obs = False
                break
        # If no parents in set, results in self.can_obs = True as it should.
        return self.can_obs

    def set_orphan(self):
        self.orphan = True

    def is_orphan(self):
        return self.orphan

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

    def set_name(self, desc):
        self.name = desc

    def get_name(self):
        return self.name

    def __lt__(self, other):
        return self.__str__() < other.__str__()

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
        self.Parents = set()
        self.orphan = False

    def get_obs_status(self, silent=False):
        return self.can_obs

    def __str__(self):
        return self.part_num

    def __repr__(self):
        return "Platform object: %s" % self.part_num


class PartGroup(object):
    def __init__(self, starting_set=set(), target_part_str=False):
        self.import_dir = os.path.join(SCRIPT_DIR, "import")
        self.Parts = starting_set

        self.target_Parts = set()
        if target_part_str:
            target_parts = target_part_str.split(",")
            # Assuming no descriptions given along w/ P/Ns through terminal.
            for target_pn in target_parts:
                self.target_Parts.add(Part(target_pn))
        else:
            self.import_target_parts()

        assert len(self.target_Parts) >= 1, "No target parts identified."

        # Use target_Parts as basis for Parts so same objects are used when
        # importing where-used reports.
        self.Parts.update(self.target_Parts)

        # Initialize list of primary parts that where-used reports pertain to.
        self.report_Parts = set()

        self.report_type = None

        print("\nParts:\t      %r" % self.Parts)
        print("Report parts: %r" % self.report_Parts)
        print("Target parts: %r" % self.target_Parts)

    def import_platforms(self, platform_dict):
        """Read in platform data from given dictionary (key is PN and value is
        True/False for can_obs).
        """
        print("\nImporting platforms...")
        for platform in platform_dict:
            platform_pn = platform.split("-")[0]
            platform_desc = platform[len(platform_pn)+1:]
            platform_obs = platform_dict[platform]
            print("\tAdding platform %12s  to group" % platform_pn)
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

    def get_target_parts(self):
        return self.target_Parts

    def print_obs_status_trace(self):
        print("")
        for PartNum in self.get_parts():
            if PartNum.__class__.__name__ == "Platform":
                print("%s is a platform" % PartNum)
            else:
                print("%s has parents %r" % (PartNum, PartNum.get_parents()))
            print("\t%s: Can OBS? %r\n" % (PartNum, PartNum.get_obs_status()))

    def get_target_obs_status(self):
        print("\nTarget parts OBS status:")
        for TargetPart in self.target_Parts:
            print("\t%s: Can OBS? %r" % (TargetPart,
                                        TargetPart.get_obs_status(silent=True)))

    def import_target_parts(self):
        """Imports all part numbers stored in import/target_parts.txt.
        Format of target_parts file can be either [P/N] or [P/N]-[DESCRIPTION].
        """
        target_filename = "target_parts.txt"
        target_parts_path = os.path.join(self.import_dir, target_filename)
        assert os.path.exists(target_parts_path), "Can't find %s" % target_filename
        print("\nImporting list of target parts from %s..." % target_filename)
        with open(target_parts_path, "r") as target_file_it:
            lines = target_file_it.read().splitlines()
            # https://stackoverflow.com/questions/19062574/read-file-into-list-and-strip-newlines
            for i, target_part in enumerate(lines):
                target_pn = target_part.split("-")[0]
                assert len(target_pn) >= 6, ("Encountered %s in file %s. "
                                            "Expected a P/N of length >= 6."
                                           % (target_part, target_filename))
                if len(target_part.split("-")) > 1:
                    # Including description isn't necessary in target_parts
                    # list.
                    target_desc = target_part[len(target_pn)+1:]
                    assert len(target_desc) > 1, ("Encountered %s in file "
                            "%s. Expected a description after P/N and dash."
                                           % (target_part, target_filename))
                else:
                    target_desc = ""
                if target_part not in self.target_Parts:
                    self.target_Parts.add(Part(target_pn, name=target_desc))
        print("...done")

    def import_all_reports(self, report_type=None):
        """Read in all where-used reports in import directory.
        """
        # Initialize list of primary parts that where-used reports pertain to.
        if report_type and self.report_type:
            raise Exception("Can't pass another report type once variable set.")
        elif report_type:
            # Should only apply first time method called.
            self.report_type = report_type
        elif self.report_type:
            # If method's already been called once w/ report type set, continue
            # using that type.
            pass
        else:
            raise Exception("Report type not specified.")

        self.report_Parts = set()

        file_list = os.listdir(self.import_dir)
        file_list.sort()

        for file_name in file_list:
            import_path = os.path.join(self.import_dir, file_name)
            if self.report_type == "SAPTC":
                self.import_SAPTC_report(import_path)
            elif self.report_type == "SAP_multi":
                self.import_SAP_multi_report(import_path)

    def import_SAPTC_report(self, import_path):
        """Read in specific where-used report from TC.
        """
        # ignore files not matching expected report pattern
        if not (os.path.basename(import_path)[:5]=="SAPTC"
                                and os.path.splitext(import_path)[-1]==".xlsx"):
            return

        print("\nReading data from %s" % os.path.basename(import_path))
        excel_data = pd.read_excel(import_path)
        import_data = pd.DataFrame(excel_data)

        part_num = import_data.iloc[0, 2]
        part_desc = import_data.iloc[1, 2]

        # Check fields are in expected locations
        assert import_data.iloc[0, 0] == "Material:", ("Expected "
                        "'Material:' in cell A2. "
                        "Check formatting in %s." % file_name)
        assert import_data.iloc[1, 0] == "Description:", ("Expected "
                        "'Description:' in cell A3. "
                        "Check formatting in %s." % file_name)

        # Rudimentary data validation
        assert len(part_num) >= 6, ("Found less than 6 digits "
                        "where part number should be in cell C2. "
                        "Check formatting in %s." % file_name)
        assert len(part_desc) > 0, ("Found empty cell where "
          "description string should be in cell C3. "
                        "Check formatting in %s." % file_name)

        # Add report part to PartsGroup
        if self.get_part(part_num) == False:
            ThisPart = Part(part_num, name=part_desc)
            ThisPart.set_report_name(os.path.basename(import_path))
            print("\tAdding %s to group (report part)" % ThisPart)
            self.add_part(ThisPart)
        else:
            print("\tPart   %s already in group (report part)" % part_num)
            ThisPart = self.get_part(part_num)
            assert ThisPart not in self.report_Parts, ("Found multiple "
                        "where-used reports in import folder for %s:\n"
                        "\t%s\n\t%s" % (ThisPart.get_pn(), ThisPart.get_report_name(),
                                                os.path.basename(import_path)))
            # If part doesn't have name/description stored, add it now.
            if not ThisPart.get_name():
                ThisPart.set_name(part_desc)
            ThisPart.set_report_name(os.path.basename(import_path))
        self.report_Parts.add(ThisPart)

        # Check table headers are in expected locations
        assert import_data.iloc[5, 3] == "Component", ("Expected "
                        "'Component' in cell D7. "
                        "Check formatting in %s." % file_name)
        assert import_data.iloc[5, 4] == "Component Description", (
            "Expected 'Component Description' in cell D7. "
                "Check formatting in %s." % file_name)

        # Iterate through the results and store in
        # if i > 6 and "End of Report" not in import_row[0]:
        for idx in import_data.index[6:-1]:
            parent_num = import_data.iloc[idx, 3]
            parent_desc = import_data.iloc[idx, 4]

            # Rudimentary data validation
            assert len(parent_num) >= 6, ("Found less than 6 digits "
                            "where part number should be in D%d. "
                            "Check formatting in %s." % (idx+2, file_name))
            assert len(parent_desc) > 0, ("Found empty cell where "
                            "description string should be in cell E%d. "
                            "Check formatting in %s." % (idx+2, file_name))

            # Create and add this part to the group if not already in
            # the Parts set.
            if self.get_part(parent_num) == False:
                NewParent = Part(parent_num, name=parent_desc)
                print("\n\tAdding %s to group" % NewParent)
                self.add_part(NewParent)
            else:
                print("\n\tPart   %s already in group" % parent_num)
                NewParent = self.get_part(parent_num)
                # If parent doesn't have name/description stored, add it now.
                if not NewParent.get_name():
                    ThisPart.set_name(parent_desc)

            # Add this part as a parent if not already in
            # the Parents set.
            if ThisPart.get_parent(parent_num) == False:
                print("\tAdding %s as parent of part %s" % (NewParent,
                                                              ThisPart))
                ThisPart.add_parent(NewParent)

        print("...done")
        print("\nParts:\t      %r" % self.Parts)
        print("Report parts: %r" % self.report_Parts)
        print("Target parts: %r" % self.target_Parts)


    def import_SAP_multi_report(self):
        pass

    def find_missing_reports(self):

        while True:
            # Every part should belong to one of these groups: parts w/ a report,
            # platforms, parts w/ "OBS" prefix, or orphan parts (empty where-used).
            platform_parts = set({Part_i for Part_i in self.Parts
                                if Part_i.__class__.__name__ == "Platform"})
            # Now exclude the platform parts from the search for parts w/ "OBS"
            obs_parts = set({Part_i for Part_i
                                        in self.Parts.difference(platform_parts)
                                                    if Part_i.get_obs_disp()})
            orphan_parts = set({Part_i for Part_i
                            in self.Parts.difference(platform_parts, obs_parts)
                                                        if Part_i.is_orphan()})

            # print("\nParts (len %d):\t      %r" % (len(self.Parts), self.Parts))
            # print("ID'd parts (len %d): %r" % (len(union_set), union_set))
            union_set = self.report_Parts.union(platform_parts, obs_parts,
                                                                  orphan_parts)
            tbd_parts = self.Parts.difference(union_set)
            if len(tbd_parts) > 0:
                print("\nMissing a report or orphan status for these parts:")
                for Part_i in tbd_parts:
                    print("\t%s" % Part_i)
                for Part_i in tbd_parts:
                    print("%s: Press Enter after adding missing report to import "
                           "or press 'n' if where-used report was empty." % Part_i)
                    answer = input("> ")
                    if answer.lower() == "n":
                        Part_i.set_orphan()
                        orphan_parts.add(Part_i)
                    else:
                        self.import_all_reports() # uses original report type.
                        break # re-generate sets
            else:
                break

    def __repr__(self):
        return "PartsGroup object: %s" % str(self.Parts)


class TreeGraph(object):
    def __init__(self, PartsGr, target_group_only=False, printout=False):
        # https://graphviz.org/doc/info/attrs.html
        # https://graphviz.org/doc/info/shapes.html
        # https://graphviz.org/doc/info/colors.html
        self.PartsGr = PartsGr
        self.target_group_only = target_group_only

        if printout:
            self.back_color = "white"
            self.part_color = "white"
        else:
            self.back_color = "slategray4"
            self.part_color = "grey"

        self.build_graph()
        # self.export_graph()

    def build_graph(self):
        self.graph = pydot.Dot(str(self.PartsGr.get_target_parts()),
                                        graph_type="graph", forcelabels=True,
                                        bgcolor=self.back_color, rankdir="TB")
        # https://stackoverflow.com/questions/19280229/graphviz-putting-a-caption-on-a-node-in-addition-to-a-label
        # https://stackoverflow.com/questions/29003465/pydot-graphviz-how-to-order-horizontally-nodes-in-a-cluster-while-the-rest-of-t

        # Create sub-graphs to enforce rank (node positioning top-to-bottom)
        self.terminal_sub = pydot.Subgraph(rank="min")
        self.target_sub = pydot.Subgraph(rank="max")
        # https://stackoverflow.com/questions/25734244/how-do-i-place-nodes-on-the-same-level-in-dot
        # https://stackoverflow.com/questions/20910596/line-up-the-heads-of-dot-graph-using-pydot?noredirect=1&lq=1

        # Create set of parts to pull from and add to graph
        if self.target_group_only:
            Parts_set = self.PartsGr.get_target_parts().copy()
        else:
            Parts_set = self.PartsGr.get_parts().copy()
        # Initialize set to hold parts already added to graph as nodes.
        self.graph_set = set()

        # Initialize incrementer to use making unique "X" nodes representing
        # no where-used results.
        inc = 0
        while len(Parts_set) > 0:
            Part_i = Parts_set.pop()
            if Part_i not in self.graph_set:
                self.add_node(Part_i)
            if Part_i.is_orphan():
                # Platforms not considered orphans
                x_node = pydot.Node("X%d" % inc, shape="box3d",
                                      style="filled", fontcolor="crimson",
                                      color="crimson", fillcolor=self.part_color,
                                      height=0.65,
                                      label="X")
                self.graph.add_node(x_node)
                self.graph.add_edge(pydot.Edge("X%d" % inc, Part_i.__str__(),
                                                            color="crimson"))
                self.terminal_sub.add_node(x_node)
                # print("Added X%d to terminal_sub" % inc)
                inc += 1

            for Parent_i in Part_i.get_parents():
                if Parent_i not in self.graph_set:
                    self.add_node(Parent_i)
                    # Add to group so its parents are included (for case where
                    # Parts_group starts out w/ only target parts)
                    Parts_set.add(Parent_i)
                if Parent_i.get_obs_status(silent=True):
                    line_color="crimson"
                else:
                    line_color="black"
                self.graph.add_edge(pydot.Edge(Parent_i.__str__(),
                                           Part_i.__str__(), color=line_color))

        self.graph.add_subgraph(self.terminal_sub)
        self.graph.add_subgraph(self.target_sub)

    def add_node(self, Part_obj):
        if Part_obj.get_obs_status(silent=True):
            outline_col = "crimson"
        else:
            outline_col = "black"

        if Part_obj.__class__.__name__ == "Platform":
            font_color = "green4"
            # print("1. %s is a platform" % Part_obj.get_pn())
        else:
            font_color = "black"

        Part_obj_node = pydot.Node(Part_obj.__str__(), shape="box3d",
                                  style="filled", fontcolor=font_color,
                                  color=outline_col, fillcolor=self.part_color,
                                  height=0.65,
                                  label="%s\n%s" %
                                  (Part_obj.get_pn(), Part_obj.get_name()))
        self.graph.add_node(Part_obj_node)
        self.graph_set.add(Part_obj)

        if Part_obj.__class__.__name__ == "Platform":
            self.terminal_sub.add_node(Part_obj_node)
            # print("Added %s to terminal_sub" % Part_obj.get_pn())
        elif Part_obj in self.PartsGr.get_target_parts():
            self.target_sub.add_node(Part_obj_node)
            # print("Added %s to target_sub" % Part_obj.get_pn())

    def export_graph(self):
        timestamp = datetime.now().strftime("%Y-%m-%dT%H%M%S")

        target_str = "+".join(map(str, sorted(self.PartsGr.get_target_parts())))
        if len(target_str) > 40:
            # If length of concatenated P/Ns exceeds 40 chars, truncate to
            # keep file name from being too long.
            target_str = "+".join(target_str[:40].split("+")[:-1]) + "..."
        export_path = os.path.join(SCRIPT_DIR, "export", "%s_%s.png"
                                                    % (timestamp, target_str))

        print("\nWriting graph to %s..." % os.path.basename(export_path))
        self.graph.write_png(export_path)
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
