print("Importing modules...")
import os
import csv
from datetime import datetime
import getpass

import pandas as pd
import numpy as np
import pydot
print("...done\n")

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory

class Part(object):
    """Object to represent a part, assy, or mod, to be used in building BOM
    structure among other Parts.
    Can add other parts to Parents set or mark as orphan.
    """
    def __init__(self, part_num, name=""):
        self.part_num = part_num
        self.name = name
        self.Parents = set()

        # Establish if part has "OBS-" prefix in SAP
        if self.name and len(self.name) > 3 and ("OBS-" in self.name[:5].upper()
                                           or "OBS -" in self.name[:6].upper()):
            self.obs_disp = True
        else:
            self.obs_disp = False

        # Initialize variable indicating if part has any parents
        self.orphan = False
        # Initialize variable indicating where-used results are present
        self.report_name = None

    def set_report_name(self, report_name):
        self.report_name = report_name

    def get_report_name(self):
        return self.report_name

    def get_report_suffix(self):
        """If part has a report, see if report has a suffix after the standard
        name and return that if so.
        """
        if not self.report_name:
            return None
        elif "SAP_multi_w" in self.report_name:
            report_prefix = "SAP_multi_w_"
        elif "SAPTC" in self.report_name:
            report_prefix = "SAPTC_BOM_Report_"
        elif "SAP_multi_BOM" in self.report_name:
            report_prefix = "SAP_multi_BOM_"
        else:
            raise Exception("Report prefix doesn't match any recognized format.")
        suffix = "_".join(os.path.splitext(self.report_name)[0].split(
                                               report_prefix)[1].split("_")[1:])
        if len(suffix) > 0:
            return suffix
        else:
            return None

    def get_obs_disp(self):
        return self.obs_disp

    def get_obs_status(self, silent=False):
        """Return True of False based on if the given part is okay to obsolete.
        First checks the obs "disposition" ("OBS" prefix in name). If obs_disp
        is False, method recursively runs on each parent, ending when either a
        platform is reached (can_obs being True or False based on platforms.py
        import) or a part has no parents (based on multi-level where-used report
        or confirmed by user).
        """
        # If the part name has OBS prefix, don't bother w/ parent query.
        if self.obs_disp:
            return True

        self.can_obs = True
        for Parent_i in self.Parents:
            # If no parents in set, leaves self.can_obs = True as it should.
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
        return self.can_obs

    def set_orphan(self):
        """Set attribute indicating part has no parents.
        """
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

    def get_parents_above(self, buffer=None):
        """Returns union of all parents above this part in the heirarchy,
        recursing up the tree.
        """
        if buffer == None:
            buffer = set()
            # This is required rather than assigning set() as the default buffer
            # in the formal parameter listing. Causes unwanted behavior.
            # https://nikos7am.com/posts/mutable-default-arguments/
        for Part_i in self.get_parents():
            if isinstance(Part_i, Platform):
                buffer.add(Part_i)
            else:
                buffer.update(set({Part_i}), Part_i.get_parents_above(buffer))

        return buffer

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
    """Represents a group of parts (Part and/or Platform objects).
    target_Parts attribute contains set of parts of interest, read from txt file.
    report_Parts attribute contains set of parts which have reports in import
    folder.
    """
    def __init__(self):
        self.import_dir = os.path.join(SCRIPT_DIR, "import")
        self.Parts = set()

        # Initialize list of parts of interest. To be populated depending on
        # program operating mode (in import_all_reports method).
        self.target_Parts = set()

        # Initialize list of parts that reports are generated for.
        self.report_Parts = set()

        # Type of report(s) being used to build PartGroup. Set in import method.
        self.report_type = None

        print("\nParts:\t      %r" % self.Parts)
        print("Report parts: %r" % self.report_Parts)
        print("Target parts: %r" % self.target_Parts)

    def import_platforms(self, platform_dict):
        """Read in platform data from given dictionary (where key is PN and
        value is True/False for can_obs).
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

    def get_parts(self, omit_platforms=False):
        if omit_platforms:
            return self.Parts - self.get_platforms()
        else:
            return self.Parts

    def get_target_parts(self):
        return self.target_Parts

    def get_report_parts(self):
        return self.report_Parts

    def get_union_bom(self):
        """Used to collect the (multi-level) BOMs of multiple parts and return
        the union of those P/Ns. The target_Parts set in this case contains the
        P/Ns whose BOMs should be union'd.
        Should be used after reading in multi-level BOM containing each target
        part.
        """
        # Ensure reports have been imported already.
        assert self.report_Parts, "No reports imported."

        if not self.target_Parts:
            self.import_target_parts(parts_update=False)

        # Ensure target parts are included in parts found in reports. Otherwise
        # will be missing its BOM members.
        if not self.target_Parts.issubset(self.Parts):
            missing_target_parts = self.target_Parts - self.target_Parts.intersection(self.Parts)
            print("\n%d of %d target parts not found in report(s):" %
                        (len(missing_target_parts), len(self.target_Parts)))
            for pn in missing_target_parts:
                print("\t%s" % pn)
            print("")
            raise Exception("%d of %d target parts not found in report(s):" %
                        (len(missing_target_parts), len(self.target_Parts)))

        # Delay adding target parts to self.Parts so above check can be conducted.
        self.Parts.update(self.target_Parts)

        # Start w/ target parts as basis for union BOM.
        union_bom = self.target_Parts.copy()
        # Test each part in group to see if the union of its parents (all the
        # way up the tree) contains any of the target parts. If so, add this
        # part to the union BOM.
        for Part_i in self.Parts:
            if Part_i in union_bom:
                continue
            elif self.target_Parts.intersection(Part_i.get_parents_above()):
                # print("Parents of %s: %r" % (NewPart, NewPart.get_parents()))
                # print("Parents above %s: %r" % (NewPart, NewPart.get_parents_above(set())))
                union_bom.add(Part_i)

        return union_bom

    def get_platforms(self):
        return set({Part_i for Part_i in self.Parts
                                               if isinstance(Part_i, Platform)})

    def print_obs_status_trace(self):
        """Print can-obsolete status for each part in Parts set.
        """
        print("")
        for PartNum in self.get_parts():
            if isinstance(PartNum, Platform):
                print("%s is a platform" % PartNum)
            else:
                print("%s has parents %r" % (PartNum, PartNum.get_parents()))
            print("\t%s: Can OBS? %r\n" % (PartNum, PartNum.get_obs_status()))

    def get_target_obs_status(self):
        """Print each part in target_Parts along with whether it is okay to OBS.
        """
        print("\nTarget parts OBS status:")
        for TargetPart in self.target_Parts:
            print("\t%s: Can OBS? %r" % (TargetPart,
                                        TargetPart.get_obs_status(silent=True)))
        if len(self.target_Parts) == 0:
            print("No target parts.")

    def import_target_parts(self, parts_update=True):
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
            for i, target_part_line in enumerate(lines):
                if not target_part_line or target_part_line.startswith("#"):
                    # Skip blank lines
                    # print("\tLine %d empty or commented out" % i)
                    continue
                target_pn = target_part_line.split("-")[0]
                assert len(target_pn) >= 6, ("Encountered %s in file %s. "
                                            "Expected a P/N of length >= 6."
                                          % (target_part_line, target_filename))
                if len(target_part_line.split("-")) > 1:
                    # Including description isn't necessary in target_parts
                    # list.
                    target_desc = target_part_line[len(target_pn)+1:]
                    assert len(target_desc) > 1, ("Encountered %s in file "
                            "%s. Expected a description after P/N and dash."
                                           % (target_part_line, target_filename))
                else:
                    target_desc = ""

                if target_pn in map(str, self.target_Parts):
                    # Make sure it's not duplicated in target_parts list
                    # May not be in self.Parts yet
                    continue
                elif self.get_part(target_pn) == False:
                    TargetPart = Part(target_pn, name=target_desc)
                else:
                    TargetPart = self.get_part(target_pn)

                self.target_Parts.add(TargetPart)
                # If target_parts file contains a duplicate, TargetPart will be
                # original object and .add() will do nothing.

        if parts_update:
            # Add target parts to overall Parts set.
            self.Parts.update(self.target_Parts)

        if len(self.target_Parts) == 0:
            print("No target parts found in %s" % target_filename)
        print("...done")

    def import_all_reports(self, report_type=None, find_missing=True,
                                                           import_subdir=None):
        """Read in all (where-used or multi-BOM) reports in import directory.
        Report type should be specified the first time this method is called.
        Subsequent calls assume same report type.
        find_missing should only be specified the first time method is called.
        """
        if report_type:
            assert report_type in ["SAPTC", "SAP_multi_w", "SAP_multi_BOM"], (
                "The only recognized report types are 'SAPTC', 'SAP_multi_w', "
                                                        "and 'SAP_multi_BOM'.")
        # Initialize list of primary parts that reports pertain to.
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

        if import_subdir:
            import_dir = os.path.join(self.import_dir, import_subdir)
        else:
            import_dir = self.import_dir

        file_list = os.listdir(import_dir)
        file_list.sort()

        if self.report_type in ["SAPTC", "SAP_multi_w"]:
            self.import_target_parts()
            assert len(self.target_Parts) >= 1, "No target parts identified."
        elif self.report_type == "SAP_multi_BOM":
            # Don't import target_Parts; not applicable when using SAP_multi_BOM
            # report type for standard case.
            # For case of creating union_bom, import_target_parts() called in
            # get_union_bom() - after reports imported.
            pass

        for file_name in file_list:
            import_path = os.path.join(import_dir, file_name)
            if self.report_type == "SAPTC":
                self.import_SAPTC_report(import_path)
            elif self.report_type == "SAP_multi_w":
                self.import_SAP_multi_w_report(import_path)
                find_missing = False
            elif self.report_type == "SAP_multi_BOM":
                self.import_SAP_multi_BOM_report(import_path)
                find_missing = False

        if find_missing:
            self.find_missing_reports()

    def import_SAPTC_report(self, import_path, verbose=False):
        """Read in a single-level where-used report generated by Teamcenter's
        SAP plugin. Create Parts objects and link parts based on BOM heirarchy.
        """
        file_name = os.path.basename(import_path)
        # ignore files not matching expected report pattern
        if not (file_name.startswith("SAPTC")
                          and os.path.splitext(file_name)[-1].lower()==".xlsx"):
            return

        print("\nReading data from %s" % file_name)
        excel_data = pd.read_excel(import_path, dtype=str)
        import_data = pd.DataFrame(excel_data)
        # https://stackoverflow.com/a/41662442

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
            ThisPart.set_report_name(file_name)
            if verbose:
                print("\tAdding %s to group (report part)" % ThisPart)
            self.add_part(ThisPart)
        else:
            if verbose:
                print("\tPart   %s already in group (report part)" % part_num)
            ThisPart = self.get_part(part_num)
            assert ThisPart not in self.report_Parts, ("Found multiple "
                       "where-used reports in import folder for %s:\n\t%s\n\t%s"
                   % (ThisPart.get_pn(), ThisPart.get_report_name(), file_name))
            # If part doesn't have name/description stored, add it now.
            if not ThisPart.get_name():
                ThisPart.set_name(part_desc)
            ThisPart.set_report_name(file_name)
        self.report_Parts.add(ThisPart)

        # Check table headers are in expected locations
        assert import_data.iloc[5, 3] == "Component", ("Expected "
                        "'Component' in cell D7. "
                        "Check formatting in %s." % file_name)
        assert import_data.iloc[5, 4] == "Component Description", (
            "Expected 'Component Description' in cell D7. "
                "Check formatting in %s." % file_name)

        # Iterate through the results and associate parent to report part.
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
                if verbose:
                    print("\n\tAdding %s to group" % NewParent)
                self.add_part(NewParent)
            else:
                if verbose:
                    print("\n\tPart   %s already in group" % parent_num)
                NewParent = self.get_part(parent_num)
                # If parent doesn't have name/description stored, add it now.
                if not NewParent.get_name():
                    NewParent.set_name(parent_desc)

            # Add this part as a parent if not already in the Parents set.
            if ThisPart.get_parent(parent_num) == False:
                if verbose:
                    print("\tAdding %s as parent of part %s" % (NewParent,
                                                              ThisPart))
                ThisPart.add_parent(NewParent)

        print("...done")
        print("\nParts:\t      %r" % self.Parts)
        print("Report parts: %r" % self.report_Parts)
        print("Target parts: %r" % self.target_Parts)


    def import_SAP_multi_w_report(self, import_path, verbose=False):
        """Read in a multi-level where-used report exported from SAP CS15.
        Create Parts objects and link parts based on BOM heirarchy.
        """
        file_name = os.path.basename(import_path)
        # ignore files not matching expected report pattern
        if not (file_name.startswith("SAP_multi_w")
                        and os.path.splitext(import_path)[-1].lower()==".xlsx"):
            return

        print("\nReading data from %s" % file_name)
        excel_data = pd.read_excel(import_path, dtype=str)
        import_data = pd.DataFrame(excel_data)
        # https://stackoverflow.com/a/41662442

        part_num = os.path.splitext(file_name)[0].split(
                                                "SAP_multi_w_")[1].split("_")[0]
        # Allowed to have additional text after P/N as long as preceded by "_".
        assert len(part_num) >= 6, ("Found less than 6 digits "
                    "where part number should be in filename (after "
                    "'SAP_multi_w_'). Check formatting of %s name." % file_name)

        # Check fields are in expected locations
        assert "Level" in import_data.columns, ("Expected "
                                        "'Level' in cell A1. "
                                        "Check formatting in %s." % file_name)
        assert "Object description" in import_data.columns, ("Expected "
                                        "'Object description' in cell D1. "
                                        "Check formatting in %s." % file_name)
        assert "Component number" in import_data.columns, ("Expected "
                                        "'Component number' in cell E1. "
                                        "Check formatting in %s." % file_name)

        # Add report part to PartsGroup
        if self.get_part(part_num) == False:
            ReportPart = Part(part_num) # no description/name given
            if verbose:
                print("\tAdding %s to group (report part)" % ReportPart)
            self.add_part(ReportPart)
        else:
            if verbose:
                print("\tPart   %s already in group (report part)" % part_num)
            ReportPart = self.get_part(part_num)
            assert ReportPart not in self.report_Parts, ("Found multiple "
                  "where-used reports in import folder for %s:\n\t%s\n\t%s"
                   % (ReportPart.get_pn(), ReportPart.get_report_name(), file_name))
        ReportPart.set_report_name(file_name)
        self.report_Parts.add(ReportPart)

        # Add extra row of NaNs to simplify loop processing.
        import_data.loc[len(import_data)] = np.nan
        # Find NaNs that divide groups.
        nan_index = import_data["Level"][import_data["Level"].isna()].index
        # If only one grouping exists, there will be no NaNs.

        # Create dictionary to store most recent part in each "level".
        level_dict = {}

        # Separate each grouping to loop through separately.
        # Identify top level of each group
        start_pos = 0
        if verbose:
            print(import_data.to_string(max_rows=10, max_cols=7))
        for break_pos in nan_index:
            max_level = int(import_data["Level"][break_pos-1])
            if verbose:
                print("\nbreak position: %d (line %d)" % (break_pos, break_pos+2))
                print("max_level: %d" % max_level)
            # Reset to base level
            ChildPart = ReportPart
            NewParent = None
            ref_level = int(import_data["Level"][start_pos])

            for i in import_data.index[start_pos:break_pos]:
                # Add parts to group as parents of earlier part.
                if verbose:
                    print("\ni: %s (line %s)" % (str(i), str(i+2)))
                parent_num = import_data["Component number"][i]
                parent_desc = import_data["Object description"][i]
                this_level = int(import_data["Level"][i])
                if verbose:
                    print("level: %d/%d" % (this_level, max_level))

                # Rudimentary data validation
                assert len(parent_num) >= 6, ("Found less than 6 digits where "
                                "part number should be in cell E%d of report. "
                                   "Check formatting in %s." % (i+2, file_name))
                assert len(parent_desc) > 0, ("Found empty cell where "
                                   "description string should be in cell D%d. "
                                   "Check formatting in %s." % (i+2, file_name))

                # Parent-setting behavior depends on if level changed since last
                # iteration.
                if this_level > ref_level:
                    # Higher level than group started at.
                    # Use previous iteration's part as the child for this part.
                    ChildPart = NewParent
                elif (NewParent and this_level < max_level
                                        and not isinstance(NewParent, Platform)):
                    # If NewParent exists (set on previous iteration), this is
                    # not first item in group.
                    # This part is at same level previous part.
                    # Keep child the same.
                    # Set previous part as orphan.
                    NewParent.set_orphan()
                elif this_level > 1:
                    # This is the first item in the group, and its level is > 1.
                    # That means it's the parent of the last part in the lower
                    # level - found in a preceding group.
                    ChildPart = level_dict[this_level-1]
                else:
                    # For first part in group, and level 1, take no action here.
                    pass
                if verbose:
                    print("ChildPart: %s" % ChildPart.__str__())

                # Create and add this part to the group if not already in
                # the Parts set.
                if self.get_part(parent_num) == False:
                    NewParent = Part(parent_num, name=parent_desc)
                    if verbose:
                        print("\tAdding %s to group" % NewParent)
                    self.add_part(NewParent)
                else:
                    if verbose:
                        print("\tPart   %s already in group" % parent_num)
                    NewParent = self.get_part(parent_num)
                    # If parent doesn't have name/description stored, add it now.
                    if not NewParent.get_name():
                        NewParent.set_name(parent_desc)
                # Store this part as most recent for current level
                level_dict[this_level] = NewParent

                # Add this part as a parent if not already in the Parents set.
                if ChildPart.get_parent(parent_num) == False:
                    if verbose:
                        print("\tAdding %s as parent of part %s" % (NewParent,
                                                                  ChildPart))
                    ChildPart.add_parent(NewParent)

                if (this_level == max_level
                                       and not isinstance(NewParent, Platform)):
                    # Everything at the highest level is either an orphan or
                    # a platform
                    if verbose:
                        print("\tSetting %s as orphan" % NewParent.__str__())
                    NewParent.set_orphan()

                # Set this level as reference level for next iteration
                ref_level = this_level

            start_pos = break_pos+1

        print("...done")
        print("\nParts:\t      %r" % self.Parts)
        print("Report parts: %r" % self.report_Parts)
        print("Target parts: %r" % self.target_Parts)


    def import_SAP_multi_BOM_report(self, import_path, verbose=False):
        """Read in a multi-level BOM exported from SAP CS12.
        Create Parts objects and link parts based on BOM heirarchy.
        """
        file_name = os.path.basename(import_path)
        # ignore files not matching expected report pattern
        if not (file_name.startswith("SAP_multi_BOM")
                        and os.path.splitext(import_path)[-1].lower()==".xlsx"):
            return

        print("\nReading data from %s" % file_name)
        excel_data = pd.read_excel(import_path, dtype=str)
        import_data = pd.DataFrame(excel_data)
        # https://stackoverflow.com/a/41662442

        part_num = os.path.splitext(file_name)[0].split(
                                            "SAP_multi_BOM_")[1].split("_")[0]
        # Allowed to have additional text after P/N as long as preceded by "_".
        assert len(part_num) >= 6, ("Found less than 6 digits "
                  "where part number should be in filename (after "
                  "'SAP_multi_BOM_'). Check formatting of %s name." % file_name)

        # Check fields are in expected locations
        assert "Explosion level" in import_data.columns, ("Expected "
                                        "'Explosion level' in cell B1. "
                                        "Check formatting in %s." % file_name)
        assert "Component number" in import_data.columns, ("Expected "
                                        "'Component number' in cell D1. "
                                        "Check formatting in %s." % file_name)
        assert "Object description" in import_data.columns, ("Expected "
                                        "'Object description' in cell E1. "
                                        "Check formatting in %s." % file_name)

        # Add report part to PartsGroup
        if self.get_part(part_num) == False:
            ReportPart = Part(part_num) # no description/name given
            if verbose:
                print("\tAdding %s to group (report part)" % ReportPart)
            self.add_part(ReportPart)
        else:
            if verbose:
                print("\tPart   %s already in group (report part)" % part_num)
            ReportPart = self.get_part(part_num)
            assert ReportPart not in self.report_Parts, ("Found multiple "
                              "reports in import folder for %s:\n\t%s\n\t%s"
               % (ReportPart.get_pn(), ReportPart.get_report_name(), file_name))
        ReportPart.set_report_name(file_name)
        self.report_Parts.add(ReportPart)

        if verbose:
            print(import_data.to_string(max_rows=10, max_cols=7))

        # Create dictionary to store most recent part in each "level".
        level_dict = {}

        Parent = ReportPart
        LastPart = Parent
        previous_level = 0
        for i in import_data.index:
            if verbose:
                print("\ni: %s (line %s)" % (str(i), str(i+2)))
            part_num = import_data["Component number"][i]
            part_desc = import_data["Object description"][i]
            if len(part_num) < 6 and part_num.startswith("CU"):
                # Skip "custom options"
                continue
            current_level = int(import_data["Explosion level"][i].split(".")[-1])
            if verbose:
                print("level: %d" % current_level)

            # Rudimentary data validation
            assert len(part_num) >= 6, ("Found less than 6 digits where "
                            "part number should be in cell D%d of report. "
                       "Check formatting in %s." % (i+2, file_name))
            assert len(part_desc) > 0, ("Found empty cell where "
                                "description string should be in cell E%d. "
                        "Check formatting in %s." % (i+2, file_name))

            if current_level > previous_level:
                if verbose:
                    print("current_level > previous_level")
                Parent = LastPart
                level_dict[previous_level] = LastPart
            elif current_level <  previous_level:
                if verbose:
                    print("current_level < previous_level")
                Parent = level_dict[current_level-1]
            else:
                # same level; keep Parent the same.
                pass

            # Create and add this part to the group if not already in
            # the Parts set.
            if self.get_part(part_num) == False:
                NewPart = Part(part_num, name=part_desc)
                if verbose:
                    print("\tAdding %s to group" % NewPart)
                self.add_part(NewPart)
            else:
                if verbose:
                    print("\tPart   %s already in group" % part_num)
                NewPart = self.get_part(part_num)
                # If parent doesn't have name/description stored, add it now.
                if not NewPart.get_name():
                    NewPart.set_name(part_desc)

            # Add this parent to this part if not already in the Parents set.
            if NewPart.get_parent(Parent.get_pn()) == False:
                if verbose:
                    print("\tAdding %s as parent of part %s" % (Parent, NewPart))
                NewPart.add_parent(Parent)

            LastPart = NewPart
            previous_level = current_level

        print("...done")
        # print("\nParts:\t      %r" % self.Parts)
        print("\nPart count:\t%d" % len(self.Parts))
        print("Report parts: %r" % self.report_Parts)
        print("Target parts: %r" % self.target_Parts)


    def find_missing_reports(self):
        """Used when importing individual where-used reports to find what reports are
        needed but not contained in import folder.
        Runs recursively after each new report is added.
        """
        while True:
            # Every part should belong to one of these groups: parts w/ a report,
            # platforms, parts w/ "OBS-" prefix, or orphan parts (empty where-used).
            platform_parts = self.get_platforms()
            # Now exclude the platform parts from the search for parts w/ "OBS-"
            obs_parts = set({Part_i for Part_i
                                        in self.Parts.difference(platform_parts)
                                                    if Part_i.get_obs_disp()})
            orphan_parts = set({Part_i for Part_i
                            in self.Parts.difference(platform_parts, obs_parts)
                                                        if Part_i.is_orphan()})

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
                        self.import_all_reports(find_missing=False)
                        # uses original report type.
                        break # re-generate sets
            else:
                break


    def export_parts_set(self, pn_set=None, omit_platforms=False):
        """Output CSV file with where-used results.
        Default is to export all parts in object. Can use pn_set to pass in the
        specific parts set desired.
        """
        timestamp = datetime.now().strftime("%Y-%m-%dT%H%M%S")
        export_path = os.path.join(SCRIPT_DIR, "export", "%s_%s_parts_set.csv"
                                  % (timestamp, self.get_pn_string(max_len=31)))

        if pn_set == None:
            parts_list = list(self.get_parts(omit_platforms))
        else:
            parts_list = list(pn_set)

        # Create new CSV file and write out.
        with open(export_path, 'w+') as output_file:
            output_file_csv = csv.writer(output_file, dialect="excel")

            print("\nWriting combined data to %s..." % os.path.basename(export_path))
            for part in parts_list:
                output_file_csv.writerow([part.get_pn(), part.get_name()])
            print("...done")


    def get_pn_string(self, pn_set_spec=False, max_len=40):
        """Generate string to represent P/N group for export filenames.
        If P/N set not specified, use target parts. If no target parts present,
        use report parts."""
        if pn_set_spec:
            pn_set = pn_set_spec.copy()
        elif self.get_target_parts():
            pn_set = self.get_target_parts().copy()
        else:
            # Cases where multi-BOM(s) used don't have target parts.
            pn_set = self.get_report_parts().copy()

        pn_str = "+".join(map(str, sorted(pn_set)))
        pn_str_suffix = ""

        if len(pn_set) == 1:
            # only one target part present; see if its report has a suffix. If
            # so, prompt for inclusion in string.
            suffix_Part = pn_set.pop()
            report_suffix = suffix_Part.get_report_suffix()

            if report_suffix:
                suffix_answer = ""
                while suffix_answer.lower() not in ["y", "n"]:
                    print("\nAppend report suffix '%s' to P/N string? [Y/N]"
                                                                % report_suffix)
                    suffix_answer = input("> ")
                if suffix_answer.lower() == "y":
                    pn_str_suffix = "_" + report_suffix
                else:
                    pass
                    # Keep it blank
            if len(pn_str_suffix) > max_len-15:
                pn_str_suffix = pn_str_suffix[:max_len-15]
                print("Truncated suffix to '%s'" % pn_str_suffix)

        if len(pn_str) > (max_len - len(pn_str_suffix)):
            # If length of concatenated P/Ns exceeds 40 chars, truncate to
            # keep file name from being too long.
            pn_str = ("+".join(pn_str[:max_len-len(pn_str_suffix)].split(
                                                             "+")[:-1]) + "...")

        return "%s%s" % (pn_str, pn_str_suffix)

    def __repr__(self):
        return "PartsGroup object: %s" % str(self.Parts)


class TreeGraph(object):
    """Object that represents a tree graph for a set of parts, showing BOM
    structure and indicating obsolete status with color.
    Graph doesn't generate correctly if run on PartGroup.target_Parts
    aren't "leaves" of the tree (base parts; can't be mods and assys).
    """
    def __init__(self, PartsGr, target_group_only=False, printout=False,
                                                             include_desc=True):
        # https://graphviz.org/doc/info/attrs.html
        # https://graphviz.org/doc/info/shapes.html
        # https://graphviz.org/doc/info/colors.html
        self.PartsGr = PartsGr
        self.target_group_only = target_group_only
        self.include_desc = include_desc

        if printout:
            self.back_color = "white"
            self.part_color = "white"
        else:
            self.back_color = "slategray4"
            self.part_color = "grey"

        self.build_graph()
        # self.export_graph()

    def build_graph(self):
        # Get username and datestamp to include on graph.
        username = getpass.getuser()
        self.timestamp = datetime.now().strftime("%Y-%m-%dT%H%M%S")
        if self.PartsGr.get_target_parts():
            part_nums = self.PartsGr.get_target_parts()
        else:
            # Cases where multi-BOM(s) used don't have target parts.
            part_nums = self.PartsGr.get_report_parts()
        self.graph = pydot.Dot(str(part_nums), graph_type="graph",
                                            forcelabels=True,
                                            bgcolor=self.back_color,
                                            rankdir="TB",
                                            label="%s %s" % (
                                                   self.timestamp.split("T")[0],
                                                                      username),
                                            labeljust="r")
        # https://stackoverflow.com/questions/19280229/graphviz-putting-a-caption-on-a-node-in-addition-to-a-label
        # https://stackoverflow.com/questions/29003465/pydot-graphviz-how-to-order-horizontally-nodes-in-a-cluster-while-the-rest-of-t

        # Create sub-graphs to enforce rank (node positioning top-to-bottom)
        self.terminal_sub = pydot.Subgraph(rank="min")
        self.target_sub = pydot.Subgraph(rank="max")
        # https://stackoverflow.com/questions/25734244/how-do-i-place-nodes-on-the-same-level-in-dot
        # https://stackoverflow.com/questions/20910596/line-up-the-heads-of-dot-graph-using-pydot?noredirect=1&lq=1

        # Create set of parts to pull from and add to graph.
        if self.target_group_only:
            Parts_set = self.PartsGr.get_target_parts().copy()
        else:
            # Omit unreferenced platforms from the graph. Needed platforms will
            # be pulled in as parents when needed.
            Parts_set = self.PartsGr.get_parts(omit_platforms=True).copy()
        # Initialize set to hold parts already added to graph as nodes.
        self.graph_set = set()

        # Initialize incrementer to use making unique "X" nodes representing
        # no where-used results.
        inc = 0
        while len(Parts_set) > 0:
            Part_i = Parts_set.pop()
            if Part_i not in self.graph_set:
                self.create_node(Part_i)
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
                inc += 1

            for Parent_i in Part_i.get_parents():
                if Parent_i not in self.graph_set:
                    self.create_node(Parent_i)
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

    def create_node(self, Part_obj):
        if Part_obj.get_obs_status(silent=True):
            outline_col = "crimson"
        else:
            outline_col = "black"

        if isinstance(Part_obj, Platform):
            font_color = "green4"
        else:
            font_color = "black"

        if self.include_desc:
            label_text = "%s\n%s" % (Part_obj.get_pn(), Part_obj.get_name())
        else:
            label_text = "%s" % Part_obj.get_pn()

        Part_obj_node = pydot.Node(Part_obj.__str__(), shape="box3d",
                                  style="filled", fontcolor=font_color,
                                  color=outline_col, fillcolor=self.part_color,
                                  height=0.65,
                                  label=label_text)
        self.graph.add_node(Part_obj_node)
        self.graph_set.add(Part_obj)

        if isinstance(Part_obj, Platform):
            self.terminal_sub.add_node(Part_obj_node)
        elif Part_obj in self.PartsGr.get_target_parts():
            self.target_sub.add_node(Part_obj_node)

    def export_graph(self):
        export_path = os.path.join(SCRIPT_DIR, "export", "%s_%s_tree.png"
                     % (self.timestamp, self.PartsGr.get_pn_string(max_len=36)))

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
