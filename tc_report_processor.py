print("Importing modules...")
import os
# import csv
from datetime import datetime

import pandas as pd
# import numpy as np
print("...done\n")

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory



def reformat_TC_single_w_report(import_path, verbose=False):
    """Read in a single-level where-used report exported from Teamcenter.
    Reformat and export after separating out superfluous results.
    """
    file_name = os.path.basename(import_path)
    assert os.path.exists(import_path), "File path not found."

    print("\nReading data from %s" % file_name)
    import_dfs = pd.read_html(import_path)
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
    return import_df

    #######################
