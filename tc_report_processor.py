print("Importing modules...")
import os
import math
from datetime import datetime
import csv
import argparse     # Used to parse optional command-line arguments
import re
import string
from colorama import Fore, Style

import pandas as pd
import numpy as np
print("...done\n")

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory

# Not all letters are available for use as revs in TC.
PROD_REV_ORDER = ["-", "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L",
                                    "M", "N", "P", "R", "T", "U", "V", "W", "Y"]
# For two-letter revs:
# Most-significant letter starts one position earlier than least-significant letter.
# Least-significant letter cycles through a list that doesn't include "".
#     "A" "B"
#  ""|___|___|
# "A"|___|___|
# "B"|___|___|
MS_REV_LETTERS = ["", "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L",
                                    "M", "N", "P", "R", "T", "U", "V", "W", "Y"]
LS_REV_LETTERS =     ["A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L",
                                    "M", "N", "P", "R", "T", "U", "V", "W", "Y"]
DISALLOWED_LETTERS = set(string.ascii_uppercase) - set(LS_REV_LETTERS)
                        # {'I', 'O', 'Q', 'S', 'X', 'Z'}

# List of columns (report fields) expected to be in TC where-used report
COL_LIST = ["Level", "Object", "Creation Date", "Current ID",
            "Current Revision", "Date Modified", "Date Released",
            "Last Modifying User", "Name", "Change", "Release Status",
            "Revisions"]


class AssumptionFail(Exception):
    pass


def print_debug(text, other_thing=None, temp=False):
    if temp:
        color_choice = Fore.BLUE
    else:
        color_choice = Fore.CYAN
    print(color_choice + "[DEBUG: %s]" % text)
    if other_thing is not None:
        print(other_thing)
    print(Style.RESET_ALL)

def is_exp_rev(rev):
    if len(rev) >= 2 and rev[-2:].isdecimal():
        # Two chars could be exp or two letters.
        # Only exp revs have three or four chars. Sorts after base letter
        return True
    else:
        # Single-char rev must be a letter or dash.
        return False

def is_prod_rev(rev):
    return not is_exp_rev(rev)

def get_latest_rev(rev_list):
    """Takes a list of revisions and returns latest rev.
    Example list: ["-", "01", "02", "A"]         - pick "A"
    Example list: ["-", "01", "02", "A", "A01"]  - pick "A"
    Example list: ["-", "01", "02", "A", "03"]   - pick "A"
    Example list: ["-", "01", "02", "03"]        - pick "-"
    Example list: ["-", "01", "02", "03", "A01"] - pick "-"
    Example list: ["-", "01", "02", "03", "03.001", "A", "A.001"] - pick "A"
    """
    # If latest sorted rev is a letter (not a letter+num), that's latest rev.
    # Want to do any check for release status here?
    # Handle if it's two letters.
    # If latest sorted rev is a letter+num, decrement and continue looking for letter (recurse).
    # If latest sorted rev is a number, check for dash rev (doesn't sort right)

    rev = rev_list[-1]
    if is_prod_rev(rev):
        return rev
    else:
        # Look for latest production rev, if one exists
        if len(rev_list) == 1:
            # One base case - searched whole list
            return rev
        # Recursive call that either finds most recent production rev or previous exp rev.
        previous_rev = get_latest_rev(rev_list[:-1])
        if is_prod_rev(previous_rev):
            return previous_rev
        else:
            # If no prod rev found, return this rev as latest.
            return rev

def extract_revs(pn, object_str):
    """Read in object list from export and extract list of revs.
    """
    # regex to pick rev out of P/N-REV-NAME:
    # "(?<=-)\-(?=-)|
    #    (?<=-)[A-HJ-MPRT-WY]{1,2}[0-9]{0,2}(?=-)|
    #        (?<=-)[A-HJ-MPRT-WY]{1,2}[0-9]{0,2}\.[0-9]{3}(?=-)|
    #            (?<=-)[0-9]{2}(?=-)|
    #                (?<=-)[0-9]{2}\.[0-9]{3}(?=-)"
    # Test cases:
    # 70663G10-A-FLOOR MAT,LWB,MTL BDY,W/HORN,HYD BRAKES
    # 652149G03---MAT,FLOOR,W/O HORN,DMND,LWB,PLASTIC BODY
    # XLWB677645G02-OBS-CHART-BB-FLOOR,MAT W/O HORN-XLWB
    # 677645G02-OBS-CHART-BB02-FLOOR,MAT W/O HORN-XLWB
    # 677645G02-OBS-CHART-B02-FLOOR,MAT W/O HORN-XLWB
    # 70663G08-V05-FLOOR MAT W/HORN,HYD BRAKES,METAL BODY
    # 70663G09-01.001-OBS-FLR MAT W/HRN,HYD BRK,MTAL BODY,DMND
    # 70663G09-B.001-OBS-FLR MAT W/HRN,HYD BRK,MTAL BODY,DMND
    # 652149G03-01-MAT,FLOOR,W/O HORN,DMND,LWB,PLASTIC BODY
    # 652149G01-02-MAT,FLOOR,LWB,DMND,W/HORN HOLES
    # 652149G01---MAT,FLOOR,LWB,DMND,W/HORN HOLES
    # 652149G01-GEOREP1---MAT,FLOOR,LWB,DMND,W/HORN HOLES
    # 652149G01-GEOREP02---MAT,FLOOR,LWB,DMND,W/HORN HOLES
    # 652149G01-GEOREP02-C-OBS-MAT,FLOOR,LWB,DMND,W/HORN HOLES
    # 677645-B-CHART-FLOOR MAT-XLWB
    # 10017652-01-10017652-BATTERY PACK UPRIGHT 6 MODULES  # P/N- in description tripped parser
    #
    # GSE specialness:
    # 739666/A-BATTERY,6CM,SAMSUNG SDI,660E-ASSY

    # Split off P/N from beginning of string
    assert object_str.startswith(pn), "extract_revs() doesn't know how to handle object_str %s \
                                that doesn't start w/ pn %s" % (object_str, pn)
    obj_str_trimmed = object_str[len(pn):]
    # Separate individual list items in rest of string. Make sure base case of one item still works.
    object_list = obj_str_trimmed.split(", %s" % pn)

    for obj in object_list:
        assert obj[0] in ("-", "/"), "extract_revs() found nonstandard delimiter \
                                                    in obj_list for pn %s" % pn

    rev_list = []
    for object in object_list:
        if object.startswith("--"):
            rev = "-"
        else:
            # Strip off leading delimiter and split off description.
            rev = object[1:].split("-")[0]
            # input("%s ->\t%s" % (object, rev)) # DEBUG
        rev_list.append(rev)
    return rev_list


def rank_rev(letter_rev):
    if letter_rev == "-":
        return -1

    ls_letter = letter_rev[-1] # least-significant
    pos_ls = LS_REV_LETTERS.index(ls_letter)

    if len(letter_rev) == 1:
        ms_letter = "" # most-significant
    if len(letter_rev) == 2:
        ms_letter = letter_rev[-2]
    pos_ms = MS_REV_LETTERS.index(ms_letter)*len(LS_REV_LETTERS)

    return pos_ls + pos_ms


def sub_bad_rev(rev, shift_fwd):
    # shift_fwd should be True or False. False indicates letter should be shifted back.
    if shift_fwd:
        plus_or_minus = lambda x, y: x + y
    else:
        plus_or_minus = lambda x, y: x - y

    # assume Z never MS letter in double-letter rev?

    # try treating disallowed letters as halfway between legit letters and round?

    if len(rev) == 1:
        # single-letter revs
        if rev == "Z" and shift_fwd:
            return "AA"
        elif rev == "Z":
            return "Y"
        else:
            new_pos = plus_or_minus(string.ascii_uppercase.index(rev), 1)
            return string.ascii_uppercase[new_pos]

    elif len(rev) == 2:
        # two-letter revs:
        if rev[0] == "Z" and shift_fwd:
            raise Exception("Don't know what to do w/ rev %s" % rev)
            # return "YY" # ?
        elif rev[0] in DISALLOWED_LETTERS:
            new_pos = plus_or_minus(string.ascii_uppercase.index(rev[0]), 1)
            if rev[1] == "Z" and shift_fwd:
                # Have to handle the case where ls "Z" would cause ms letter to
                # increment again below.
                rev = string.ascii_uppercase[new_pos] + "A"
            else:
                rev = string.ascii_uppercase[new_pos] + rev[1]

        if rev[1] == "Z" and shift_fwd:
            new_ms_pos = plus_or_minus(rank_rev(rev[0]), 1)
            return LS_REV_LETTERS[new_ms_pos] + "A"
        elif rev[1] in DISALLOWED_LETTERS:
            new_pos = plus_or_minus(string.ascii_uppercase.index(rev[1]), 1)
            rev = rev[0] + string.ascii_uppercase[new_pos]

        return rev

    else:
        raise Exception("sub_bad_rev() received input of more than two letters"
                                                                ": %s" % rev)


def get_rev_difference(rev, newer_rev):
    if is_exp_rev(newer_rev) or is_exp_rev(rev):
        return False
    elif rev == newer_rev:
        # Need this explicit to avoid error if disallowed letters included.
        return 0
    elif len(newer_rev) > 2 or len(rev) > 2:
        raise Exception("get_rev_difference() called with nonstandard revs:\n"
                                            "\t%s -> %s" % (rev, newer_rev))
    elif (len(newer_rev) == 2 and len(rev) == 2) and (rev[0] == newer_rev[0]):
        # Same most-sig letter. Just compare least-sig letters.
        return get_rev_difference(rev[1], newer_rev[1])
    else:
        # Handle legacy revs using now-disallowed letters.
        # Promote or demote to closest allowed neighbor
        if set([*rev]).intersection(DISALLOWED_LETTERS):
            # Determine if single letter or either of two letters are disallowed.
            rev = sub_bad_rev(rev, shift_fwd=False)
        if set([*newer_rev]).intersection(DISALLOWED_LETTERS):
            newer_rev = sub_bad_rev(newer_rev, shift_fwd=True)

        return rank_rev(newer_rev) - rank_rev(rev)


def two_rev_diff(rev, newer_rev):
    return ( (get_rev_difference(rev, newer_rev) != False)
         and (get_rev_difference(rev, newer_rev) > 1) )

def parse_rev_status(status_str):
    # "Concept"                             (check mark)
    # "Concept Cancelled"                   (check mark w/ red slash) - deployed 2024-01-30. Developed in 2023-11-16 meeting.
    # "Baseline"                            (check mark)
    # "Alpha"                               (check mark)
    # "Beta"                                (check mark)
    # "Gamma"                               (check mark)
    # "Gamma,Concept"                       (check mark) - e.g. 668404-03
    # "Concept,Approved"                    (green flag)
    # "Alpha,Approved"                      (green flag)
    # "Beta,Approved"                       (green flag)
    # "Gamma,Approved"                      (green flag)
    exp_status = re.findall(r"(concept|baseline|alpha|beta|gamma)$", status_str,
                                                            flags=re.IGNORECASE)
    grn_status = re.findall(r"(,approved)$", status_str,
                                                            flags=re.IGNORECASE)
    purple_status = re.findall(r"(preliminary)$", status_str,
                                                            flags=re.IGNORECASE)

    canc_status = re.findall(r"(concept cancelled)$", status_str,
                                                            flags=re.IGNORECASE)

    turf_checkd = re.findall(r"(engineering_released|ppap_release|engrework|quarantined|obsoleted|voided)$", status_str,
                                                            flags=re.IGNORECASE)

    # "Engineering Released"                (yellow flag)
    yel_status = re.findall(r"(engineering released)$", status_str,
                                                            flags=re.IGNORECASE)
    # "Engineering Released -Superseded"    (yellow flag - strikethrough)
    sup_status = re.findall(r"(-superseded)$", status_str, flags=re.IGNORECASE)

    # "Engineering Released,Released"       (checkered flag)
    # "Released"                            (checkered flag)
    # "Engineering Released,Redline Release"(red checkered flag)
    # "Redline Release"                     (red checkered flag) - not sure this exists
    checkd_status = re.findall(r"((?<!engineering.)released)$", status_str,
                                                            flags=re.IGNORECASE)
    rcheckd_status = re.findall(r"(redline release)$", status_str,
                                                            flags=re.IGNORECASE)


    # "Overtaken"                            (checkered flag w/ red dash sign)
    ovtkn_status = re.findall(r"(overtaken)$", status_str, flags=re.IGNORECASE)

    # "Obsolete"                            (red X)
    obs_status = re.findall(r"(obsolete)$", status_str, flags=re.IGNORECASE)

    if sum([len(exp_status),
           len(grn_status),
           len(purple_status),
           len(canc_status),
           len(turf_checkd),
           len(yel_status),
           len(sup_status),
           len(checkd_status),
           len(rcheckd_status),
           len(ovtkn_status),
           len(obs_status)]) > 1:
        raise Exception("More than one status match found: %s" % status_str)

    if not status_str:
        # If empty string, then rev isn't statused at all.
        return "unstatused"
    elif len(exp_status) == 1:
        return "exp_statused"
    elif len(canc_status) == 1:
        return "canc_status"
    elif len(grn_status) == 1:
        return "green_flag"
    elif len(purple_status) == 1:
        return "purple_flag"
    elif len(turf_checkd) == 1:
        return "checkered_flag_other"
    elif len(yel_status) == 1:
        return "yellow_flag"
    elif len(sup_status) == 1:
        return "superseded_yellow"
    elif len(checkd_status) == 1:
        return "checkered_flag"
    elif len(rcheckd_status) == 1:
        return "red_checkered_flag"
    elif len(ovtkn_status) == 1:
        return "overtaken"
    elif len(obs_status) == 1:
        return "obsolete"
    else:
        raise Exception("No valid status found in '%s'" % status_str)

def convert_date(date_str):
    if date_str:
        timestamp = datetime.strptime(date_str, "%d-%b-%Y %H:%M")
        return datetime.strftime(timestamp, "%Y-%m-%d")
    else:
        return ""

def parse_report_pn(report_name, base_only=False):
    """Setting base_only argument to True will strip any "GEOREP"
    """
    # Report name format ex.:
    #   "2022-03-10_637381-GEOREP1--_TC_where-used.html"
    #   "2022-02-02_614575-A_TC_where-used.html"

    if not report_name.endswith("TC_where-used.html"):
        return False
    report_date = report_name.split("_")[0]
    if report_name.split("_")[1][-2:] == "--":
        report_rev = "-"
    else:
        report_rev = report_name.split("_")[1].split("-")[-1]
    report_pn = report_name.split(report_date + "_")[1].split("-" + report_rev)[0]

    if base_only:
        # if not "-GEOREP" in report_name.upper():
        #     raise Exception("base_only arg specified in parse_report_pn() but "
        #                             "'-GEOREP' not found in report filename.")
        report_pn = report_pn.upper().split("-GEOREP")[0]
    # report_pn_match = re.findall(r"(?<=^\d{4}-\d{2}-\d{2}_)[\w-]+(?=-[\w-]_TC_where-used)", report_name, flags=re.IGNORECASE)
    # if len(report_pn_match) == 1:
    #     self.report_pn = report_pn_match

    return report_pn


class TCReport(object):
    """Object representing single TC where-used report.
    """
    def __init__(self, import_path):
        self.file_path = os.path.abspath(import_path)
        assert os.path.exists(self.file_path), "File path not found."

        self.file_name = os.path.basename(self.file_path)

    def get_core_df(self):
        return self.core_df

    def get_import_df(self):
        return self.import_df

    def get_extra_df(self):
        return self.extra_df

    def get_filename(self):
        return self.file_name

    def import_report(self, verbose=False):
        """Read in a single-level where-used report exported from Teamcenter.
        Returns dataframe with table data.
        """
        # Extract report part number from file name
        self.report_pn = parse_report_pn(self.file_name)
        assert self.report_pn != False, "TCReport.import_report() failed. \
                      Filename not recognized as TC report: %s" % self.file_name

        print("Reading data from %s" % self.file_name)
        import_dfs = pd.read_html(self.file_path)
        # Returns list of dfs. List should only have one df.
        assert len(import_dfs) == 1, "Irregular HTML table format found."
        self.import_df = import_dfs[0]
        # Have to manually set dtype of some cols to string in case report
        # contains only numeric P/Ns.
        self.import_df["Current ID"] = self.import_df["Current ID"].astype(str)
        # Also must z-fill single-digit exp revs if they've been represented as integers.
        self.import_df["Current Revision"] = self.import_df["Current Revision"].astype(str).apply(lambda x: x.zfill(2) if x.isdecimal() else x)

        for col in COL_LIST:
            assert col in self.import_df.columns, ("Column '%s' not found in "
                 "TC report '%s'.\nExpecting these columns: \n"
                                        % (col, self.file_name) + str(COL_LIST))

        if verbose:
            print(self.import_df.loc[:, ["Current ID", "Current Revision", "Name"]])

        lev_0_result = self.import_df[self.import_df["Level"] == 0]["Current ID"].values[0]
        assert lev_0_result == self.report_pn, \
            "P/N in report name doesn't match level-0 result in report table.\n%s\n%s" \
            % (self.report_pn, lev_0_result)
        print("...done\n")

    def reformat_dataframe(self, verbose=False):
        """Reformat single-level where-used report data and separate out
        superfluous results. Put primary columns of interest at left, renamed.
        Keep original report columns at right.
        """
        # Get rid of report part from table.
        base_df = self.import_df.drop(self.import_df[self.import_df["Level"]==0].index)
        # https://pythoninoffice.com/delete-rows-from-dataframe/

        # Sort so P/Ns are grouped, and revs within those groups are sorted.
        base_df.sort_values(by=["Current ID", "Current Revision"], inplace=True)
        # https://datatofish.com/sort-pandas-dataframe/
        # Probably the same:
        # base_df.sort_values(by=["Object"], inplace=True)

        # Reset index after sorting
        # Original index values get saved to new col called "index".
        base_df.reset_index(inplace=True)
        base_df.rename(columns={"index": "Original Row Num [DEBUG]"}, inplace=True)
        # https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.reset_index.html

        # Create new column for comments, to be used to explain filtering.
        base_df["Comments"] = ""

        # Create new column w/ list of revs extracted from "Revisions" string.
        # See extract_revs() function defined above.
        base_df["Rev List [DEBUG]"] = base_df.apply(
                        lambda x: extract_revs(x["Current ID"], x["Revisions"]),
                                                                        axis=1)
        # https://stackoverflow.com/questions/34279378/python-pandas-apply-function-with-two-arguments-to-columns

        # Create new column w/ report P/N so each row can be traced back to
        # original report if multiple reports combined (like case of GEOREPs).
        base_df["Report P/N [DEBUG]"] = self.report_pn

        # Create new column w/ latest rev extracted from "Revisions" string.
        # See get_latest_rev() function defined above.
        base_df["Latest Rev"] = base_df["Rev List [DEBUG]"].apply(get_latest_rev)

        # Create new column w/ status of this P/N-rev combo.
        # Blank status fields will be nan, so replace these w/ empty strings first.
        base_df["Rev Status [DEBUG]"] = base_df["Release Status"].fillna("").apply(parse_rev_status)

        base_df["Last Mod Date"] = base_df["Date Modified"].fillna("").apply(convert_date)

        # Duplicate most pertinent columns and arrange at left. Original report
        # columns will be hidden in export.
        renamed_cols = ["Part Number", "Revision", "Name (Teamcenter)"] # new names
        base_df[renamed_cols] = base_df[["Current ID", "Current Revision", "Name"]]

        # Rename original cols to indicate name mapping like "Current ID [=> "Part Number"]"
        base_df.rename(columns={"Current ID": "Current ID [=> \"Part Number\"]"},
                                                                inplace=True)
        base_df.rename(columns={"Name": "Name [=> \"Name (Teamcenter)\"]"},
                                                                inplace=True)
        base_df.rename(columns={"Current Revision": "Current Revision [=> \"Revision\"]"},
                                                                inplace=True)
        base_df.rename(columns={"Date Modified": "Date Modified [=> \"Last Mod Date\"]"},
                                                                inplace=True)

        # Position specific columns at beginning, including ones previously created.
        # Leaves all original report columns at right.
        first_cols = renamed_cols + ["Latest Rev", "Last Mod Date", "Comments", \
           "Rev Status [DEBUG]", "Rev List [DEBUG]", "Report P/N [DEBUG]", \
                                                    "Original Row Num [DEBUG]"]
        base_df = base_df[first_cols + [col for col in \
                                      base_df.columns if col not in first_cols]]
        # https://stackoverflow.com/questions/44009896/python-pandas-copy-columns

        # Handle rare case of duplicate P/N-rev results (e.g. 605563-F report)
        # Have to do this before splitting base_df below.
        base_df["PN-Rev"] = base_df[["Part Number", "Revision"]].agg('-'.join, axis=1) # temporary col
        # https://stackoverflow.com/questions/19377969/combine-two-columns-of-text-in-pandas-dataframe
        # https://stackoverflow.com/questions/32918506/pandas-how-to-filter-dataframe-for-duplicate-items-that-occur-at-least-n-times
        for pn_rev in set(base_df["PN-Rev"].fillna("")):
            # Select all rows w/ this P/N-rev combo.
            pn_rev_filter = base_df["PN-Rev"]==pn_rev
            pn_rev_rows = base_df[pn_rev_filter]

            # Find cases where P/N-rev combo is duplicated.
            if len(pn_rev_rows) > 1:
                # print_debug("dups:", other_thing=pn_rev_rows[["PN-Rev", "Release Status"]])
                # print_debug("choice:", other_thing=pn_rev_rows[~pn_rev_rows["Release Status"].isna()][["PN-Rev", "Release Status"]])
                # Seems in each case, only one of the group has a status.
                no_release_status_filter = pn_rev_rows["Release Status"].isna()
                no_change_filter = pn_rev_rows["Change"].isna() # alternate criterion
                # Check assumption
                valid_count_rlst = len(pn_rev_rows[~no_release_status_filter]) # should be 1
                valid_count_chg = len(pn_rev_rows[~no_change_filter]) # should be 1
                if valid_count_rlst == 1:
                    discard_filter = no_release_status_filter
                elif valid_count_chg == 1:
                    discard_filter = no_change_filter
                else:
                    raise AssumptionFail("Unsure how to choose which result "
                            "duplicate is correct (expected only one row to "
                             "have non-empty Release Status or Change field)",
                            pn_rev_rows[["PN-Rev", "Change", "Release Status"]])
                # Drop duplicate rows w/o status
                to_be_dropped = base_df.loc[pn_rev_rows[discard_filter].index]
                # print_debug("to be dropped:", other_thing=to_be_dropped[["PN-Rev", "Release Status"]])
                base_df.drop(index=to_be_dropped.index, inplace=True)
                # print_debug("dups after dropping:", other_thing=base_df[base_df["PN-Rev"]==pn_rev][["PN-Rev", "Release Status"]])
        # Remove temporary PN-Rev col.
        base_df.drop(columns=["PN-Rev"], inplace=True)

        # Build filters to move study files, exp revs, etc. to another dataframe
        # that will be appended to end of export.
        georep_filter = base_df["Part Number"].str.upper().str.contains("GEOREP")
        # Remove items where P/N starts w/ letter.
        letter_pn_filter = ~base_df["Part Number"].str[:3].str.isdecimal()
        chart_name_filter = base_df["Name (Teamcenter)"].str.upper().str.startswith("CHART")
        study_name_filter = base_df["Name (Teamcenter)"].str.upper().str.contains("STUDY")
        study_pn_filter = base_df["Part Number"].str.upper().str.contains("STUDY")
        # Sub empty string in place of NaNs in base_df["Release Status"] for eval (not inplace).
        obs_status_filter = base_df["Release Status"].fillna("").str.contains("Obsolete")

        # Add comments to help user interpret results.
        base_df.loc[georep_filter, "Comments"] = "GEOREP [grey highlight]"
        base_df.loc[letter_pn_filter, "Comments"] = "Part number starting with letters"
        base_df.loc[chart_name_filter, "Comments"] = "Chart drawing"
        base_df.loc[study_name_filter, "Comments"] = "Study file [grey highlight]"
        base_df.loc[study_pn_filter, "Comments"] = "Study file [grey highlight]"
        base_df.loc[obs_status_filter, "Comments"] = "Obsolete status [red highlight]"
        # Some old-rev comments will be overwritten below when checking for
        # newer rev in report.
        # https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy


        extra_filter = (   georep_filter    |  letter_pn_filter |
                          chart_name_filter | study_name_filter |
                            study_pn_filter |     obs_status_filter   )
        # https://stackoverflow.com/a/54030143
        # https://datagy.io/python-isdigit/

        self.extra_df = pd.DataFrame(columns=base_df.columns)

        # Move extranneous rows to extra_df.
        self.extra_df = pd.concat([self.extra_df, base_df[extra_filter]])
        self.core_df = base_df.drop(base_df[extra_filter].index)

        # Isolate latest rev of each thing. Not necessarily thing that sorts last.
        pn_set = set(self.core_df["Part Number"].fillna(""))
        for pn in pn_set:
            # Select all rows w/ this P/N.
            pn_filter = self.core_df["Part Number"]==pn
            pn_rows = self.core_df[pn_filter]

            report_rev_list = list(self.core_df.loc[pn_filter, "Revision"]) # Only revs included in report.
            latest_rev_in_report = get_latest_rev(report_rev_list) # Not necessarily latest rev in TC
            # Every "Latest Rev" value in pn_rows is the same, so use first one.
            latest_rev_glob = self.core_df.loc[pn_filter, "Latest Rev"].iloc[0]

            # Identify various types of "old" revs (not mutually exclusive)
            # global: at least one newer prod rev exists in TC
            # two: more than one newer prod rev exists in TC
            # exp: a prod rev exists in TC, whereas this rev is exp.
            # rep: a newer (prod or exp) rev exists in this report
            old_revs_glob_filter = (pn_filter) & (self.core_df["Revision"]!=latest_rev_glob)

            old_revs_two_filter = (pn_filter) & (self.core_df["Revision"].apply(
                                            lambda x: two_rev_diff(x, latest_rev_glob)))
            # print("\npn: %s\n" % pn) # DEBUG
            old_revs_rep_filter = (pn_filter) & (self.core_df["Revision"]!=latest_rev_in_report)

            old_revs_exp_filter = ( (pn_filter) &
                                    (self.core_df["Revision"].apply(is_exp_rev)) &
                                    (is_prod_rev(latest_rev_glob))                    )

            # Move all but the latest rev in report to extra_df
            # Move latest rev in report too if more than one newer production rev exists in TC.
            # Revs to move from core_df to extra_df:
            move_filter = (old_revs_two_filter | old_revs_rep_filter | old_revs_exp_filter)

            # Apply commments in order to layer over each other.
            self.core_df.loc[old_revs_glob_filter, "Comments"] = "Newer rev exists [yellow highlight]"
            self.core_df.loc[old_revs_two_filter, "Comments"] = "Newer statused rev in TC [yellow highlight]"
            self.core_df.loc[old_revs_exp_filter, "Comments"] = "Production rev exists [yellow highlight]"
            self.core_df.loc[old_revs_rep_filter, "Comments"] = "Newer rev in report [yellow highlight]"

            self.extra_df = pd.concat([self.extra_df, self.core_df[move_filter]]).sort_index()
            self.core_df.drop(self.core_df[move_filter].index, inplace=True)


class TCReportGroup(object):
    def __init__(self, dir_path):
        self.report_dir = dir_path

    def find_reports(self, pn=False, single_report_path=False):
        """Pass either P/N or specific report path, but not both.
        Searches self.report_dir for reports associated w/ given base P/N.
        """
        assert pn != single_report_path, "find_reports() requires " \
                            "either P/N or specific report path, but not both."
        self.report_set = set()

        if single_report_path:
            self.add_report(TCReport(single_report_path))
            self.base_pn = parse_report_pn(os.path.basename(single_report_path),
                                                                base_only=True)
        elif pn:
            self.base_pn = pn
            for item in sorted(os.listdir(self.report_dir)):
                item_path = os.path.join(self.report_dir, item)
                if not os.path.isfile(item_path) or not item.endswith("TC_where-used.html"):
                    continue
                elif parse_report_pn(item, base_only=True) == self.base_pn:
                    report_path = os.path.join(self.report_dir, item)
                    self.add_report(TCReport(report_path))
                    # parse_report_pn re-run in TCReport.__init__() w/o base_only

            if len(self.report_set) < 1:
                raise Exception("Found no reports matching P/N %s in %s." % (self.base_pn, self.report_dir))
        else:
            raise Exception("find_reports() requires one arg of either P/N "
                                                    "or specific report path")

    def add_report(self, Report):
        self.report_set.add(Report)

    def process_reports(self):
        for Report in self.report_set:
            Report.import_report()
            Report.reformat_dataframe()

    def combine_reports(self):
        # Gather all core_dfs and extra_dfs
        core_dfs = [Report.get_core_df() for Report in self.report_set]
        extra_dfs = [Report.get_extra_df() for Report in self.report_set]

        # Combine all core_dfs and export_dfs
        core_df_combo = pd.concat(core_dfs).sort_values(by=["Object"])
        # Eliminate duplicate P/N-rev combos (if base P/N and GEOREP are both used in same file)
        core_df_combo.drop_duplicates(subset="Object", inplace=True)
        # print_debug("core_df_combo:", other_thing=core_df_combo[["Part Number", "Revision", "Latest Rev"]])

        extra_df_combo = pd.concat(extra_dfs).sort_values(by=["Object"])
        extra_df_combo.drop_duplicates(subset="Object", inplace=True)
        # print_debug("extra_df_combo:", other_thing=extra_df_combo[["Part Number", "Revision", "Latest Rev"]])

        # Combine core_df_combo and extra_df_combo rows w/ 4 blank rows in between.
        buffer_df = pd.DataFrame(np.nan, index=range(0, 4), columns=extra_df_combo.columns)
        self.export_df = pd.concat([core_df_combo, buffer_df, extra_df_combo], ignore_index=True)

    def export(self):
        """Output XLSX file with reordered and combined data from report(s).
        """
        self.combine_reports()

        if len(self.report_set) > 1:
            combined = "s"
        else:
            combined = ""

        timestamp = datetime.now().strftime("%Y-%m-%dT%H%M%S")
        export_path = os.path.join(self.report_dir,
                                        "%s_%s_processed_TC_report%s.xlsx" %
                                        (timestamp, self.base_pn, combined))

        print("Writing combined data to %s..." % os.path.basename(export_path))
        with pd.ExcelWriter(export_path, engine="xlsxwriter") as writer:
            # Reorder and select columns
            sheet1 = "TC_%s" % self.base_pn
            self.export_df.to_excel(writer, sheet_name=sheet1, index=False,
                                                            freeze_panes=(1,1))

            # Format spreadsheet
            # https://xlsxwriter.readthedocs.io/working_with_pandas.html
            # https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html
            # https://xlsxwriter.readthedocs.io/worksheet.html
            # https://xlsxwriter.readthedocs.io/format.html
            workbook = writer.book
            worksheet = writer.sheets[sheet1]

            # Left-justify format (applied below)
            l_align = workbook.add_format()
            l_align.set_align('left')

            # Right-justify format (applied below)
            r_align = workbook.add_format()
            r_align.set_align('right')

            # Center-justify format (applied below)
            c_align = workbook.add_format()
            c_align.set_align('center')

            # Store column numbers and letters, and cell ranges
            pn_col_num = self.export_df.columns.get_loc("Part Number")
            rev_col_num = self.export_df.columns.get_loc("Revision")
            name_col_num = self.export_df.columns.get_loc("Name (Teamcenter)")
            latestrev_col_num = self.export_df.columns.get_loc("Latest Rev")
            mdate_col_num = self.export_df.columns.get_loc("Last Mod Date")
            comments_col_num = self.export_df.columns.get_loc("Comments")
            status_col_num = self.export_df.columns.get_loc("Rev Status [DEBUG]")
            revlist_col_num = self.export_df.columns.get_loc("Rev List [DEBUG]")
            reportpn_col_num = self.export_df.columns.get_loc("Report P/N [DEBUG]")
            origrownum_col_num = self.export_df.columns.get_loc("Original Row Num [DEBUG]")

            pn_col_letter = string.ascii_uppercase[pn_col_num]
            rev_col_letter = string.ascii_uppercase[rev_col_num]
            name_col_letter = string.ascii_uppercase[name_col_num]
            latestrev_col_letter = string.ascii_uppercase[latestrev_col_num]
            status_col_letter = string.ascii_uppercase[status_col_num]
            # https://stackoverflow.com/questions/4528982/convert-alphabet-letters-to-number-in-python#4528997

            pn_cell_range = '%s2:%s10000' % (pn_col_letter, pn_col_letter)
            rev_cell_range = '%s2:%s10000' % (rev_col_letter, rev_col_letter)
            name_cell_range = '%s2:%s10000' % (name_col_letter, name_col_letter)
            latestrev_cell_range = '%s2:%s10000' % (latestrev_col_letter, latestrev_col_letter)

            # Specify column widths and justifications
            worksheet.set_column(pn_col_num, pn_col_num, 16, l_align)
            worksheet.set_column(rev_col_num, rev_col_num, 8, l_align)
            worksheet.set_column(name_col_num, name_col_num, 45, l_align)
            worksheet.set_column(latestrev_col_num, latestrev_col_num, 9, c_align)
            worksheet.set_column(mdate_col_num, mdate_col_num, 13, l_align)
            worksheet.set_column(comments_col_num, comments_col_num, 35, l_align)

            # Hide DEBUG columns
            # Make col width equal char count of heading
            worksheet.set_column(status_col_num, status_col_num,
                              len("Rev Status [DEBUG]"), None, {"hidden": True})
            worksheet.set_column(revlist_col_num, revlist_col_num,
                                len("Rev List [DEBUG]"), None, {"hidden": True})
            # Add extra width to this one for autofilter button (added below)
            worksheet.set_column(reportpn_col_num, reportpn_col_num,
                            len("Report P/N [DEBUG]")+4, None, {"hidden": True})
            worksheet.set_column(origrownum_col_num, origrownum_col_num,
                        len("Original Row Num [DEBUG]"), None, {"hidden": True})

            # Add filter button
            # https://xlsxwriter.readthedocs.io/working_with_autofilters.html
            # https://xlsxwriter.readthedocs.io/working_with_cell_notation.html#cell-notation

            worksheet.autofilter(0, reportpn_col_num, 10000, reportpn_col_num)
            #              row_start,   col_start,   row_end,    col_end
            # Not possible to selectively filter discontinuous ranges.

            # Collapse original report columns carried over from TC export.
            # Most users will probably not care about these, but keeping them for
            # ref and debugging.
            col_num_start = origrownum_col_num + 1
            col_num_end = len(self.export_df.columns) - 1
            worksheet.set_column(col_num_start, col_num_end, None, None,
                                                   {"level": 1, "hidden": True})
            worksheet.set_column(col_num_end+1, col_num_end+1, None, None,
                                                            {"collapsed": True})
            # https://xlsxwriter.readthedocs.io/working_with_outlines.html

            # Set header row ht
            # worksheet.set_row(0, 30)
            # Set up wrap text for header. Having to do this somewhat manually.
            # header_format = workbook.add_format({'bold': True})
            # header_format.set_align('center')
            # header_format.set_align('bottom')
            # worksheet.set_row(0, header_format)
            # worksheet.write('A1', "Current ID", header_format)
            # worksheet.write('B1', "Revision", header_format)
            # worksheet.write('C1', "Name", header_format)

            # Light red fill with dark red text.
            red_hl_ft = workbook.add_format({'bg_color':   '#FFC7CE',
                                             'font_color': '#9C0006' })
            # Highlight OBS
            worksheet.conditional_format(name_cell_range,
                                {'type': 'text', 'criteria': 'begins with',
                                                            'value': 'OBS',
                                                        'format': red_hl_ft})

            worksheet.conditional_format(name_cell_range,
                        {'type': 'formula',
                         'criteria': '=$%s2="obsolete"' % status_col_letter,
                         'format': red_hl_ft})

            # Light yellow fill with dark yellow text.
            yellow_hl_ft = workbook.add_format({'bg_color':   '#FFEB9C',
                                                'font_color': '#9C6500'})
            worksheet.conditional_format(latestrev_cell_range,
                                    {'type': 'formula', 'criteria': '=$D2<>$B2',
                                                       'format': yellow_hl_ft})

            # Grey fill.
            grey_hl = workbook.add_format({'bg_color': '#CFCFCF'})
            # Grey out study files
            worksheet.conditional_format(name_cell_range,
                                    {'type': 'text', 'criteria': 'begins with',
                                                            'value': 'STUDY',
                                                            'format': grey_hl})
            worksheet.conditional_format(name_cell_range,
                                    {'type': 'text', 'criteria': 'begins with',
                                                            'value': 'Study',
                                                            'format': grey_hl})
            worksheet.conditional_format(name_cell_range,
                                    {'type': 'text', 'criteria': 'begins with',
                                                            'value': 'study',
                                                            'format': grey_hl})
            worksheet.conditional_format(pn_cell_range,
                                    {'type': 'text', 'criteria': 'containing',
                                                            'value': 'STUDY',
                                                            'format': grey_hl})
            worksheet.conditional_format(pn_cell_range,
                                    {'type': 'text', 'criteria': 'containing',
                                                            'value': 'Study',
                                                            'format': grey_hl})
            worksheet.conditional_format(pn_cell_range,
                                    {'type': 'text', 'criteria': 'containing',
                                                            'value': 'study',
                                                            'format': grey_hl})

            # Grey out P/Ns that start w/ letters
            worksheet.conditional_format(pn_cell_range,
            {'type': 'formula',
             'criteria': '=NOT(IFERROR(IF(ISBLANK($%s2),TRUE,(INT(LEFT($%s2,3)))), FALSE))'
             % (pn_col_letter, pn_col_letter),
             'format': grey_hl})

            # Grey out GEOREPs.
            worksheet.conditional_format(pn_cell_range,
                                    {'type': 'text', 'criteria': 'containing',
                                                            'value': 'GEOREP',
                                                            'format': grey_hl})

            # Grey out exp revs.
            worksheet.conditional_format(rev_cell_range, {'type': 'formula',
             'criteria': '=AND(NOT(ISBLANK($%s2)),ISNUMBER(INT(RIGHT($%s2,1))))'
             % (rev_col_letter, rev_col_letter),
             'format': grey_hl})

            # Green fill.
            # green_hl = workbook.add_format({'bg_color':   '#92D050'})

            # Add status images to Revision column.
            for n, status in enumerate(self.export_df["Rev Status [DEBUG]"].fillna("")):
                row_num = n+1
                if is_exp_rev(str(self.export_df.loc[n, "Revision"])):
                     img_name = "%s_greybg.png" % status
                else:
                    img_name = "%s.png" % status
                img_relpath = "./img/%s" % img_name
                img_abspath = "%s/%s" % (SCRIPT_DIR, img_relpath)
                if not status or status == "unstatused":
                    pass
                elif not os.path.exists(img_abspath):
                    print("\tWarning: Missing status image: %s" % img_relpath)
                    pass
                else:
                    worksheet.insert_image(row_num, rev_col_num, img_abspath,
                                                {'x_offset': 40, 'y_offset': 2})

            # Get rid of green triangles in output sheet.
            worksheet.ignore_errors({"number_stored_as_text": "A2:Z10000"})
            # https://xlsxwriter.readthedocs.io/worksheet.html#ignore_errors

            print("...done")

def convert_win_path(path_str):
    """Converts Windows path to Linux path."""
    drive_letter = path_str[0]
    return path_str.replace("\\", "/").replace("%s:" % drive_letter,
                                            "/mnt/%s" % drive_letter.lower())


if __name__ == "__main__":
    # Don't run if module being imported. Only if script being run directly.
    parser = argparse.ArgumentParser(description="Program to process TC "
                                                        "where-used reports")
    parser.add_argument("-f", "--file", help="Specify path to TC report to "
                                "import for processing", type=str, default=None)
    parser.add_argument("-d", "--dir", help="Specify dir containing TC reports "
                            "to import for processing", type=str, default=None)
    parser.add_argument("-p", "--pn", help="Specify part num that reports in "
           "dir (specified w/ --dir option) pertain to", type=str, default=None)
    # https://www.programcreek.com/python/example/748/argparse.ArgumentParser
    args = parser.parse_args()

    if args.file:
        path_str = convert_win_path(args.file)
        assert os.path.isfile(path_str), "Not a valid file path: %s" % args.file
        assert not args.dir, "Can only pass file or dir, not both."
        ReportGroup = TCReportGroup(os.path.dirname(path_str))
        ReportGroup.find_reports(single_report_path=path_str)
    elif args.dir:
        path_str = convert_win_path(args.dir)
        assert os.path.isdir(path_str), "Not a valid directory path: %s" % args.dir
        if not args.pn:
            pn = None
            # Read in all report P/Ns in the dir and see if only one P/N found.
            dir_contents = os.listdir(path_str)
            for item in dir_contents:
                base_pn = parse_report_pn(os.path.join(path_str, item), base_only=True)
                if base_pn == False:
                    # Not a TC report. Skip
                    continue

                if pn is None:
                    pn = base_pn
                elif pn == base_pn:
                    # Keep checking P/Ns
                    continue
                else:
                    # Runs if pn was previously set, but now program sees a different P/N.
                    pn = None
                    break

            # print("DEBUG: pn: %s" % pn)
            # This runs only if above loop found more than one base P/N:
            while not pn:
                print("Enter base part num:")
                pn = input("> ").upper()
        else:
            pn = args.pn
        ReportGroup = TCReportGroup(path_str)
        ReportGroup.find_reports(pn=pn)
    else:
        raise Exception("Need to pass TC report file path or dir path.")

    ReportGroup.process_reports()
    ReportGroup.export()


# # Reference
# import_df.mask(~extra_filter)
# # This returns df w/ NaNs other than things ID'd by filter.
# import_df[extra_filter]
# # This returns df w/ only things ID'd by filter.
# import_df.index[import_df["Current ID"].str.upper().str.contains("STUDY")]
# # This returns indices of rows in import_df where "STUDY" items are
# import_df.drop(import_df.index[import_df["Current ID"].str.upper().str.contains("STUDY")])
# import_df.drop(import_df[extra_filter].index)
# # This returns version of import_df w/ "STUDY" items removed
# import_df[import_df["Current ID"].str.upper().str.contains("STUDY")]
# # This returns "STUDY" rows of import_df.
