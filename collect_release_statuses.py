import os, shutil
import argparse
from tabulate import tabulate

import tc_report_processor as tcr_proc



def extract_release_statuses(Report, stat_dict=None):
    """Takes in tcr_proc.TCReport object and extracts all release statuses.
    Optional dict argument allows statuses to be added to an existing status
    dict. Otherwise, new dict will be used (and not returned).
    """
    Report.import_report()
    df = Report.get_import_df()
    # print(df[["Current ID", "Current Revision", "Release Status"]]) # DEBUG

    if stat_dict is None:
        stat_dict = {}

    for n, status_str in enumerate(df["Release Status"].fillna("")):
        if status_str == "":
            continue
        elif stat_dict.get(status_str):
            continue
        else:
            stat_dict[status_str] = ("%s-%s" % (df["Current ID"][n],
                                                df["Current Revision"][n]),
                                                          Report.get_filename())


def collect_statuses(dir_path):
    """Reads in every html TC report found in dir_path, extracts release statuses
    found in each, and collects them in one dictionary.
    """
    status_dict = {}

    # Read in all report P/Ns in the dir.
    path_str = tcr_proc.convert_win_path(dir_path)
    dir_contents = os.listdir(path_str)
    for item in dir_contents:
        pn = tcr_proc.parse_report_pn(os.path.join(path_str, item), base_only=False)
        if pn == False:
            # Not a TC report.
            continue

        ThisReport = tcr_proc.TCReport(os.path.join(path_str, item))

        try:
            extract_release_statuses(ThisReport, status_dict) # adds to dict
        except (AssertionError, KeyError):
            # Occurs if old report encountered that doesn't include all columns.
            print("\tFailed - %s" % os.path.basename(ThisReport.get_filename()))
            # Just skip this report
            continue


    print(tabulate([[key, status_dict[key][0], status_dict[key][1]]
                                        for key in sorted(status_dict)],
                                        headers=["Status", "P/N-Rev", "File"]))


if __name__ == "__main__":
    # Don't run if module being imported. Only if script being run directly.
    parser = argparse.ArgumentParser(description="Program to aggregate all release "
                                    "statuses found in TC where-used reports")
    parser.add_argument("-d", "--dir", help="Specify dir containing TC reports "
                            "to extract status strings from", type=str, default=None)

    args = parser.parse_args()
    if args.dir:
        path_str = convert_win_path(args.dir)
        collect_statuses(path_str)

    else:
        raise Exception("Must pass directory path containing TC reports.")