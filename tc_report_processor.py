print("Importing modules...")
import os
import math
from datetime import datetime
import csv
import argparse     # Used to parse optional command-line arguments

import pandas as pd
import numpy as np
print("...done\n")

# dir path where this script is stored
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# https://stackoverflow.com/questions/29768937/return-the-file-path-of-the-file-not-the-current-directory

# Not all letters are available for use as revs in TC.
PROD_REV_ORDER = ["-", "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L",
                                    "M", "N", "P", "R", "T", "U", "V", "W", "Y"]
# List of columns (report fields) expected to be in TC where-used report
COL_LIST = ["Level", "Object", "Creation Date", "Current ID",
            "Current Revision", "Date Modified", "Date Released",
            "Last Modifying User", "Name", "Change", "Release Status",
            "Revisions"]

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
    object_list = object_str.split(pn + "-")
    rev_list = []
    for object in object_list:
        if object.startswith("--"):
            rev = "-"
        else:
            # Standard case requires splitting off dash and name.
            # If part has no name, this leaves a ", " at end of each rev, so
            # split that off. Has no effect on standard parts.
            rev = object.split("-")[0].split(", ")[0]
        rev_list.append(rev)
    return rev_list[1:] # First item in list is ''


class TCReport(object):
    """Object representing single TC where-used report.
    """
    def __init__(self, import_path):
        self.file_path = os.path.abspath(import_path)
        assert os.path.exists(self.file_path), "File path not found."

        self.file_name = os.path.basename(self.file_path)
        self.dir_path = os.path.dirname(self.file_path)

    def import_report(self, verbose=False):
        """Read in a single-level where-used report exported from Teamcenter.
        Returns dataframe with table data.
        """
        # Report name format ex.:
        #   "2022-03-10_637381-GEOREP1--_TC_where-used.html"
        #   "2022-02-02_614575-A_TC_where-used.html"
        report_date = self.file_name.split("_")[0]
        if self.file_name.split("_")[1][-2:] == "--":
            report_rev = "-"
        else:
            report_rev = self.file_name.split("_")[1].split("-")[-1]
        self.report_pn = self.file_name.split(report_date + "_")[1].split("-" + report_rev)[0]

        print("Reading data from %s" % self.file_name)
        import_dfs = pd.read_html(self.file_path)
        # Returns list of dfs. List should only have one df.
        assert len(import_dfs) == 1, "Irregular HTML table format found."
        self.import_df = import_dfs[0]

        for col in COL_LIST:
            assert col in self.import_df.columns, ("Column '%s' not found in "
                 "TC report. Expecting these columns: \n" % col + str(COL_LIST))

        if verbose:
            print(self.import_df.loc[:, ["Current ID", "Current Revision", "Name"]])

        assert self.import_df[self.import_df["Level"] == 0]["Current ID"].values[0] == self.report_pn, \
                "P/N in report name doesn't match level-0 result in report table."
        print("...done\n")

    def reformat_report(self, verbose=False):
        """Read in a single-level where-used report exported from Teamcenter.
        Reformat and separate out superfluous results.
        Put primary columns of interest at left, renamed.
        Keep original report columns at right.
        """
        # Get rid of report part from table.
        core_df = self.import_df.drop(self.import_df[self.import_df["Level"]==0].index)
        # https://pythoninoffice.com/delete-rows-from-dataframe/

        # Sort so P/Ns are grouped, and revs within those groups are sorted.
        core_df.sort_values(by=["Current ID", "Current Revision"], inplace=True)
        # https://datatofish.com/sort-pandas-dataframe/
        # Is this any different?
        # core_df.sort_values(by=["Object"], inplace=True)

        # Create new column for comments, to be used to explain filtering.
        core_df["Comments"] = ""

        # Create new column w/ list of revs extracted from "Revisions" string.
        # See extract_revs() function defined above.
        core_df["Rev List [DEBUG]"] = core_df.apply(
                        lambda x: extract_revs(x["Current ID"], x["Revisions"]),
                                                                        axis=1)
        # https://stackoverflow.com/questions/34279378/python-pandas-apply-function-with-two-arguments-to-columns

        # Create new column w/ report P/N so each row can be traced back to
        # original report if multiple reports combined (like case of GEOREPs).
        core_df["Report P/N [DEBUG]"] = self.report_pn

        # Create new column w/ latest rev extracted from "Revisions" string.
        # See get_latest_rev() function defined above.
        core_df["Latest Rev"] = core_df["Rev List [DEBUG]"].apply(get_latest_rev)

        # Build filters to move study files, exp revs, etc. to another dataframe
        # that will be appended to end of export.
        # Remove items where P/N starts w/ letter.
        letter_pn_filter = ~core_df["Current ID"].str[:3].str.isdecimal()
        chart_name_filter = core_df["Name"].str.upper().str.startswith("CHART")
        study_name_filter = core_df["Name"].str.upper().str.contains("STUDY")
        study_pn_filter = core_df["Current ID"].str.upper().str.contains("STUDY")
        # Sub empty string inplace of NaNs in core_df["Release Status"] for eval (not inplace).
        obs_pn_filter = core_df["Release Status"].fillna("").str.contains("Obsolete")

        # Add comments to help user interpret results.
        core_df.loc[letter_pn_filter, "Comments"] = "Part number starting with letters"
        core_df.loc[chart_name_filter, "Comments"] = "Chart drawing"
        core_df.loc[study_name_filter, "Comments"] = "Study file [grey highlight]"
        core_df.loc[study_pn_filter, "Comments"] = "Study file [grey highlight]"
        core_df.loc[obs_pn_filter, "Comments"] = "Obsolete status [red highlight]"
        # https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy


        extra_filter = (   letter_pn_filter | chart_name_filter |
                          study_name_filter |   study_pn_filter | obs_pn_filter )
        # https://stackoverflow.com/a/54030143
        # https://datagy.io/python-isdigit/

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
        self.export_df = core_df.append(buffer_df.append(extra_df, ignore_index=True), ignore_index=True)

        # Duplicate most pertinent columns and arrange at left. Original report
        # columns will be hidden in export.
        renamed_cols = ["Part Number", "Revision", "Name (Teamcenter)"] # new names
        self.export_df[renamed_cols] = self.export_df[["Current ID", "Current Revision", "Name"]]

        # Position specific columns at beginning, including ones previously created.
        # Leaves all original report columns at right.
        first_cols = renamed_cols + ["Latest Rev", "Comments", \
                                       "Rev List [DEBUG]", "Report P/N [DEBUG]"]
        self.export_df = self.export_df[first_cols + [col for col in \
                               self.export_df.columns if col not in first_cols]]
        # https://stackoverflow.com/questions/44009896/python-pandas-copy-columns

        # Rename original cols to indicate name mapping like "Current ID [=> "Part Number"]"
        self.export_df.rename(columns={"Current ID": "Current ID [=> \"Part Number\"]"},
                                                                inplace=True)
        self.export_df.rename(columns={"Name": "Name [=> \"Name (Teamcenter)\"]"},
                                                                inplace=True)
        self.export_df.rename(columns={"Current Revision": "Current Revision [=> \"Revision\"]"},
                                                                inplace=True)


    def export_report(self):
        """Output CSV file with reordered rows from original report.
        """
        timestamp = datetime.now().strftime("%Y-%m-%dT%H%M%S")
        export_path = os.path.join(self.dir_path,
                        "%s_%s_processed_TC_report.xlsx" % (timestamp, self.report_pn))

        print("Writing combined data to %s..." % os.path.basename(export_path))
        with pd.ExcelWriter(export_path, engine="xlsxwriter") as writer:
            # Reorder and select columns
            sheet1 = "TC_%s" % self.report_pn
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

            # Specify column widths and justifications
            col_num = self.export_df.columns.get_loc("Part Number")
            worksheet.set_column(col_num, col_num, 20, r_align)

            col_num = self.export_df.columns.get_loc("Revision")
            worksheet.set_column(col_num, col_num, 8, c_align)

            col_num = self.export_df.columns.get_loc("Name (Teamcenter)")
            worksheet.set_column(col_num, col_num, 45)

            col_num = self.export_df.columns.get_loc("Latest Rev")
            worksheet.set_column(col_num, col_num, 9, r_align)

            col_num = self.export_df.columns.get_loc("Comments")
            worksheet.set_column(col_num, col_num, 35, l_align)

            # Hide "Rev List [DEBUG]" and "Report P/N [DEBUG]" columns
            # Make col width equal char count of heading
            col_num = self.export_df.columns.get_loc("Rev List [DEBUG]")
            worksheet.set_column(col_num, col_num, len("Rev List [DEBUG]"),
                                                        None, {"hidden": True})
            col_num = self.export_df.columns.get_loc("Report P/N [DEBUG]")
            worksheet.set_column(col_num, col_num, len("Report P/N [DEBUG]"),
                                                        None, {"hidden": True})

            # Collapse original report columns carried over from TC export.
            # Most users will probably not care about these, but keeping them for
            # ref and debugging.
            col_num_start = self.export_df.columns.get_loc("Report P/N [DEBUG]") + 1
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
                                            'font_color': '#9C0006'})
            # Highlight OBS
            worksheet.conditional_format('C2:C10000', {'type': 'text',
                                                      'criteria': 'begins with',
                                                      'value': 'OBS',
                                                      'format': red_hl_ft})
            worksheet.conditional_format('C2:C10000', {'type': 'formula',
                                                       'criteria': '=$R2="Obsolete"',
                                                       'format': red_hl_ft})

            # Light yellow fill with dark yellow text.
            yellow_hl_ft = workbook.add_format({'bg_color':   '#FFEB9C',
                                                'font_color': '#9C6500'})

            # Highlight cases where latest rev newer than rev found by where-used
            worksheet.conditional_format('B2:B10000', {'type': 'formula',
                                                       'criteria': '=$D2<>$B2',
                                                       'format': yellow_hl_ft})
            worksheet.conditional_format('D2:D10000', {'type': 'formula',
                                                       'criteria': '=$D2<>$B2',
                                                       'format': yellow_hl_ft})

            # Grey fill.
            grey_hl = workbook.add_format({'bg_color':   '#BFBFBF'})
            # Grey out study files
            worksheet.conditional_format('C2:C10000', {'type': 'text',
                                                  'criteria': 'begins with',
                                                  'value': 'STUDY',
                                                  'format': grey_hl})
            worksheet.conditional_format('C2:C10000', {'type': 'text',
                                               'criteria': 'begins with',
                                               'value': 'Study',
                                               'format': grey_hl})
            worksheet.conditional_format('C2:C10000', {'type': 'text',
                                           'criteria': 'begins with',
                                           'value': 'study',
                                           'format': grey_hl})
            worksheet.conditional_format('A2:A10000', {'type': 'text',
                                                 'criteria': 'containing',
                                                 'value': 'STUDY',
                                                 'format': grey_hl})
            worksheet.conditional_format('A2:A10000', {'type': 'text',
                                              'criteria': 'containing',
                                              'value': 'Study',
                                              'format': grey_hl})
            worksheet.conditional_format('A2:A10000', {'type': 'text',
                                          'criteria': 'containing',
                                          'value': 'study',
                                          'format': grey_hl})

            # Grey out P/Ns that start w/ letters
            worksheet.conditional_format('A2:A10000', {'type': 'formula',
             'criteria': '=NOT(IFERROR(IF(ISBLANK($A2),TRUE,(INT(LEFT($A2,3)))), FALSE))',
             'format': grey_hl})

            # Green fill.
            # Highlight production revs green
            green_hl = workbook.add_format({'bg_color':   '#92D050'})
            worksheet.conditional_format('B2:B10000', {'type': 'formula',
               'criteria': '=NOT(IFERROR(IF(ISBLANK($B2),TRUE,(INT(RIGHT($B2,2)))), FALSE))',
               'format': green_hl})

            print("...done")


def run(import_path):
    Report = TCReport(import_path)
    Report.import_report()
    Report.reformat_report()
    Report.export_report()



parser = argparse.ArgumentParser(description="Program to process TC where-used"
                                                                    " reports")
parser.add_argument("-f", "--file", help="Specify path to TC report to import "
                    "for processing", type=str, default=None)
# https://www.programcreek.com/python/example/748/argparse.ArgumentParser
args = parser.parse_args()

assert args.file, "Need to pass TC report file path."

drive_letter = args.file[0]
import_path = args.file.replace("\\", "/").replace("%s:" % drive_letter,
                                                "/mnt/%s" % drive_letter.lower())

run(import_path)


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
