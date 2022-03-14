print("Importing modules...")
import os
import math
from datetime import datetime

import pandas as pd
# import numpy as np
print("...done\n")

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory

# Not all letters are available for use as revs in TC.
PROD_REV_ORDER = ["-", "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L",
                                    "M", "N", "P", "R", "T", "U", "V", "W", "Y"]

def is_exp_rev(rev):
    if len(rev) >= 2:
        # Two chars could be exp or two letters.
        # Only exp revs have three or four chars. Sorts after base letter
        try:
            rev_num = int(rev[-2:])
            return True
        except ValueError:
            return False
    else:
        # Single-char rev must be a letter or dash.
        return False

def is_prod_rev(rev):
    return not is_exp_rev(rev)

def get_latest_rev(rev_list):
    """Takes a list of revisions and returns latest rev.
    """
    #   If latest sorted rev is a letter (not a letter+num), that's latest rev. Want to do any check for release status here?
    #   Handle if it's two letters.
    #   If latest sorted rev is a letter+num, decrement and continue looking for letter (recurse).
    #   If latest sorted rev is a number, check for dash rev (doesn't sort right)

    # Example list: ["-", "01", "02", "A"] - pick "A"
    # Example list: ["-", "01", "02", "A", "A01"] - pick "A"
    # Example list: ["-", "01", "02", "A", "03"] - pick "A"
    # Example list: ["-", "01", "02", "03"] - pick "-"
    # Example list: ["-", "01", "02", "03", "A01"] - pick "-"

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


def reformat_TC_single_w_report(import_path, verbose=False):
    """Read in a single-level where-used report exported from Teamcenter.
    Reformat and export after separating out superfluous results.
    """
    file_path = os.path.realpath(import_path)
    assert os.path.exists(file_path), "File path not found."

    file_name = os.path.basename(file_path)
    print("\nReading data from %s" % file_name)
    import_dfs = pd.read_html(file_path)
    print("...done")
    # Returns list of dfs. List should only have one df.
    assert len(import_dfs) == 1, "Irregular HTML table format found."
    import_df = import_dfs[0]
    if verbose:
        print(import_df.loc[:, ["Current ID", "Current Revision", "Name"]])
    # Get rid of report part from table
    import_df.drop(import_df[import_df["Level"]==0].index, inplace=True)
    # https://pythoninoffice.com/delete-rows-from-dataframe/

    # Sort so P/Ns are grouped, and revs within those groups are sorted.
    import_df.sort_values(by=["Current ID", "Current Revision"], inplace=True)
    # https://datatofish.com/sort-pandas-dataframe/
    # Is this any different?
    # import_df.sort_values(by=["Object"], inplace=True)
    return import_df

    #######################
