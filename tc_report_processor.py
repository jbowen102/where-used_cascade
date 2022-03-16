print("Importing modules...")
import os
import math
from datetime import datetime
import csv

import pandas as pd
import numpy as np
print("...done\n")

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory

# Not all letters are available for use as revs in TC.
PROD_REV_ORDER = ["-", "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L",
                                    "M", "N", "P", "R", "T", "U", "V", "W", "Y"]

def is_exp_rev(rev):
    if len(rev) >= 2 and rev[-2:].isnumeric():
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
    """
    #   If latest sorted rev is a letter (not a letter+num), that's latest rev. Want to do any check for release status here?
    #   Handle if it's two letters.
    #   If latest sorted rev is a letter+num, decrement and continue looking for letter (recurse).
    #   If latest sorted rev is a number, check for dash rev (doesn't sort right)

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

def extract_revs(object_str):
    """Read in object list from export and extract list of revs.
    """
    object_list = object_str.split(", ")
    rev_list = []
    for object in object_list:
        if "---" in object:
            rev = "-"
        else:
            rev = object.split("-")[1]
        rev_list.append(rev)
    return rev_list


def import_TC_single_w_report(import_path, verbose=False):
    """Read in a single-level where-used report exported from Teamcenter.
    Returns dataframe with table data.
    """
    file_path = os.path.realpath(import_path)
    assert os.path.exists(file_path), "File path not found."
    file_name = os.path.basename(file_path)

    # Report name format ex.:
    #   "2022-03-10_637381-GEOREP1--_TC_where-used.html"
    #   "2022-02-02_614575-A_TC_where-used.html"
    report_date = file_name.split("_")[0]
    report_pn = file_name.split("_")[1][:-2] # remove rev and dash from end.

    print("Reading data from %s" % file_name)
    import_dfs = pd.read_html(file_path)
    # Returns list of dfs. List should only have one df.
    assert len(import_dfs) == 1, "Irregular HTML table format found."
    import_df = import_dfs[0]
    if verbose:
        print(import_df.loc[:, ["Current ID", "Current Revision", "Name"]])

    assert import_df[import_df["Level"] == 0]["Current ID"].values[0] == report_pn, \
            "P/N in report name doesn't match level-0 result in report table."
    print("...done\n")
    return import_df


def reformat_TC_single_w_report(report_df, verbose=False):
    """Read in a single-level where-used report exported from Teamcenter.
    Reformat and export after separating out superfluous results.
    """
    # Get rid of report part from table.
    core_df = report_df.drop(report_df[report_df["Level"]==0].index)
    # https://pythoninoffice.com/delete-rows-from-dataframe/

    # Sort so P/Ns are grouped, and revs within those groups are sorted.
    core_df.sort_values(by=["Current ID", "Current Revision"], inplace=True)
    # https://datatofish.com/sort-pandas-dataframe/
    # Is this any different?
    # core_df.sort_values(by=["Object"], inplace=True)

    # Build filters to move study files, exp revs, etc. to another dataframe
    # that will be appended to end of export.
    # Remove items where P/N starts w/ letter.
    extra_filter = ((core_df["Current ID"].str.upper().str.contains("STUDY"))
                  | (core_df["Name"].str.upper().str.startswith("STUDY"))
                  | (core_df["Name"].str.upper().str.startswith("CHART"))    )
    # https://stackoverflow.com/a/54030143

    extra_df = pd.DataFrame(columns=core_df.columns)

    # Move extranneous rows to extra_df.
    extra_df = extra_df.append(core_df[extra_filter])
    # https://stackoverflow.com/questions/15819050/pandas-dataframe-concat-vs-append
    core_df.drop(core_df[extra_filter].index, inplace=True)

    # Isolate latest rev of each thing. Not necessarily thing that sorts last.
    for id in core_df["Current ID"]:
        # For each line, ID the P/N. Select all rows w/ this P/N.
        id_rows = core_df[core_df["Current ID"]==id]
        rev_list = list(id_rows["Current Revision"])
        latest_rev = get_latest_rev(rev_list)

        # Move all but the latest rev to extra_df
        old_revs = id_rows[id_rows["Current Revision"]!=latest_rev]
        extra_df = extra_df.append(old_revs)
        core_df.drop(old_revs.index, inplace=True)

    # Combine core and extra rows w/ 4 blank rows in between.
    buffer_df = pd.DataFrame(np.nan, index=range(0, 4), columns=extra_df.columns)
    export_df = core_df.append(buffer_df.append(extra_df, ignore_index=True), ignore_index=True)

    return export_df


def export_report(export_df, report_pn):
    """Output CSV file with reordered rows from original report.
    """
    timestamp = datetime.now().strftime("%Y-%m-%dT%H%M%S")
    export_path = os.path.join(SCRIPT_DIR, "export",
                    "%s_%s_processed_TC_report.xlsx" % (timestamp, report_pn))

    print("Writing combined data to %s..." % os.path.basename(export_path))
    # Reorder and select columns
    export_df.to_excel(export_path, index=False, freeze_panes=(1,1),
                            columns=["Current ID", "Current Revision", "Name"])
    # Can col width be set?

    print("...done")


#######################
import_path = "./reference/2022-03-14_630034--_TC_where-used.html"
report_df = import_TC_single_w_report(import_path)
output_df = reformat_TC_single_w_report(report_df)
export_report(output_df, "630034")

#######################

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
