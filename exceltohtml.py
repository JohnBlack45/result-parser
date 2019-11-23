from xlrd import open_workbook, xldate_as_tuple, xldate_as_datetime
from tabulate import tabulate
from datetime import time
import pathlib
import os
import hashlib
from operator import itemgetter

HTML_HEADER = \
"""<!DOCTYPE html>
<html lang="en">
<head>
  <title>Bootstrap Example</title>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css">
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js"></script>
</head>
<body>             
<div class="container">"""

CSS_SCRIPT = \
"""
<style type="text/css">
  th { background-color: #074d32; color: white; font-size:12pt; }
</style>
"""

HTML_FOOTER = \
"""</div>
</body>
</html>
"""

HEADER_STRINGS = {
    "position": ["pos", "pl.", "place"],
    "first_name": ["firstname", "forename", "name1"],
    "surname": ["lastname", "surname", "name2"],
    "full_name": ["name", "athlete"],
    "club": ["club", "team"],
    "category": ["cat", "age", "group", "category", "gender"],
    "chip_time": ["chip", "net"],
    "finish_time": ["gun", "finish", "time", "gross"],
    "bib": ["bib", "num"],
    "lap": ["lap"],
    "distance": ["half", "km", "mile"]
}

WHEELCHAIR_ATHLETES = ["jim corbett", "paul hannan", "karol doherty", "james divin"]

class ExcelDateFormatException(Exception):
    pass

def strip_excessive_spaces(input_string):
    input_string = input_string.lstrip().rstrip()
    input_string = " ".join(input_string.split())
    return input_string

class Runner():
    def __init__(self, name):

        self.name = strip_excessive_spaces(name).lower().title()
        self.first_name = "Empty"
        self.surname = "Empty"
        self.wins = 0

        split_name = self.name.split(" ")
        if len(split_name) == 2:
            self.first_name = split_name[0]
            self.surname = split_name[1]

    def get_full_name(self):
        if self.first_name is "Empty" or self.surname is "Empty":
            return self.name
        else:
            return f"{self.first_name} {self.surname}"

    def convert_to_list(self):
        return [self.get_full_name(), self.wins]

    def get_key(self):
        if self.first_name is "Empty" or self.surname is "Empty":
            return self.name
        else:
            sorted_list = sorted([self.first_name, self.surname])
            return f"{sorted_list[0]} {sorted_list[1]}"

    def __repr__(self):
        return self.get_full_name()


class Sheet:

    def __init__(self, parent, index):
        self.parent = parent
        self.workbook = parent.workbook
        self.handle = self.workbook.sheet_by_index(index)
        self.index = index
        self.html_table = ""
        self.is_empty = self.is_empty_sheet()
        self.heading_row = -1
        self.winner = ""
        self.columns = {}

        global info

        if not self.is_empty:
            self.heading_row = self.find_heading_row()

            if self.heading_row != -1:
                info.headings_found += 1
                self.identify_columns()

                runner = Runner(self.get_winner())
                self.winner = runner
                if runner.get_key() in info.winners.keys():
                    info.winners[runner.get_key()].wins += 1
                else:
                    runner.wins = 1
                    info.winners[runner.get_key()] = runner
            else:
                info.headings_not_found += 1

    def build_html_table(self):

        global info

        table = []

        starting_row = 0
        if self.heading_row:
            starting_row = self.heading_row

        for row in range(starting_row, self.handle.nrows):
            row_list = []

            if self.is_row_empty(row):
                continue

            for col in range(0, self.handle.ncols):
                cell = self.handle.cell(row, col)
                value = self.handle.cell(row, col).value

                # Float = 2
                if cell.ctype == 2:

                    if value.is_integer():
                        value = int(value)

                # Date = 3
                elif cell.ctype == 3:
                    try:
                        time_raw = xldate_as_tuple(cell.value, self.workbook.datemode)
                        value = str(time(*time_raw[3:]))
                    except Exception as e:
                        info.date_format_issue += 1
                        raise ExcelDateFormatException

                row_list.append(value)

            table.append(row_list)

        # Generate the html for the table
        self.html_table = tabulate(table, tablefmt='html', headers="firstrow")
        self.bootstrap_html_table()
        return self.html_table

    def is_empty_sheet(self):

        if self.handle.nrows < 10:
            self.is_empty = True
            return self.is_empty

        num_rows_to_check = 15
        if self.handle.nrows < num_rows_to_check:
            num_rows_to_check = self.handle.nrows

        for row in range(0, num_rows_to_check):
            for col in range(0, self.handle.ncols):
                cell_value = self.handle.cell(row, col).value
                if cell_value:
                    self.is_empty = False
                    return self.is_empty

        self.is_empty = True
        return self.is_empty

    def find_heading_row(self):

        num_rows_to_check = 20

        if self.handle.nrows < num_rows_to_check:
            num_rows_to_check = self.handle.nrows

        for row in range(0, num_rows_to_check):

            heading_matches_in_row = 0

            for col in range(0, self.handle.ncols):
                cell_string = str(self.handle.cell(row, col).value).lower().replace(" ", "")

                for heading in HEADER_STRINGS.values():
                    for pattern in heading:

                        if pattern == cell_string:
                            heading_matches_in_row += 1

                            if heading_matches_in_row > 2:
                                return row
        return -1

    def is_row_empty(self, row, count=0):

        value_found = 0

        for col in range(0, self.handle.ncols):
            cell_value = self.handle.cell(row, col).value
            if cell_value:
                if value_found == count:
                    return False
                else:
                    value_found += 1
        return True

    def identify_columns(self):

        for col in range(0, self.handle.ncols):

            cell_value = str(self.handle.cell(self.heading_row, col).value).lower().replace(" ", "")

            for heading in HEADER_STRINGS.keys():
                match_strings = HEADER_STRINGS.get(heading)
                for string in match_strings:
                    if string == cell_value:
                        if heading not in self.columns:
                            self.columns[heading] = col
                        break

    def get_winner(self):

        first_result_row_max_search = self.heading_row + 6
        first_result_row = first_result_row_max_search

        # Find the first result row after the header this may be a few rows down
        for row in range(self.heading_row + 1, first_result_row_max_search):
            if self.is_row_empty(row, 3):
                continue
            else:

                winner = ""

                if "full_name" in self.columns:
                    winner = str(self.handle.cell(row, self.columns["full_name"]).value)
                elif "first_name" in self.columns and "surname" in self.columns:
                    first_name = str(self.handle.cell(row, self.columns["first_name"]).value)
                    surname = str(self.handle.cell(row, self.columns["surname"]).value)
                    winner = f"{first_name} {surname}"

                if winner.lower() in WHEELCHAIR_ATHLETES:
                    print(f"Wheelchair Found {self.parent.path}")
                    winner = ""
                    first_result_row_max_search += 1
                    continue

                first_result_row = row
                break

        if first_result_row == first_result_row_max_search:
            #return f"No winner found - no result found in first x rows - {self.parent.path}"
            return ""

        winner = winner.replace("\"", "")
        winner = winner.replace(",", "")
        winner = winner.replace("*", "") # ? Mad i know
        winner = winner.replace(",", "")
        winner = winner.replace("’", "")
        winner = winner.replace("'", "")
        winner = winner.replace("O'", "O")
        winner = winner.replace("o'", "O")
        winner = winner.replace("Mc ", "Mc")
        winner = winner.replace("mc ", "Mc")

        return winner


    def bootstrap_html_table(self):
        self.html_table = self.html_table.replace("<table>", "<table class=\"table table-striped table-hover table-sm\">")


class ExcelFile:

    def __init__(self, path):
        self.path = path
        self.html = ""
        self.heading_row = ""
        self.sheets = []
        self.workbook = open_workbook(path)

    def build_html(self):

        global info

        self.html = HTML_HEADER + CSS_SCRIPT

        for sheet_index in range(0, self.workbook.nsheets):

            info.num_sheets += 1

            sheet = Sheet(self, sheet_index)
            if sheet.is_empty:
                info.empty_sheets += 1
                continue

            # Add a blank line between every sheet
            if sheet_index != 0:
                self.html += "<br>"

            #print(f"{self.path} {sheet.winner}")

            #try:
            #    self.html += sheet.build_html_table()
            #except ExcelDateFormatException:
            #    return ""

        self.html += HTML_FOOTER


def sha256sum(filename):
    h = hashlib.sha256()
    b = bytearray(128*1024)
    mv = memoryview(b)
    with open(filename, 'rb', buffering=0) as f:
        for n in iter(lambda : f.readinto(mv), 0):
            h.update(mv[:n])
    return h.hexdigest()


RESULT_DIRECTORY = "output/"
UNSUPPORTED_FILE_TYPES = [".pdf", ".doc", ".docx"]


class Info:

    def __init__(self):
        self.num_files = 0
        self.unsupported_files = []
        self.supported_file = []
        self.file_hashes = set()
        self.headings_found = 0
        self.headings_not_found = 0
        self.num_sheets = 0
        self.empty_sheets = 0
        self.date_format_issue = 0
        self.duplicate_files_found = 0
        self.exception_files = 0
        self.winners = {}

    def __repr__(self):
        table = [["# Files", self.num_files],
                ["Supported", len(self.supported_file)],
                ["Unsupported", len(self.unsupported_files)],
                ["Duplicates", self.duplicate_files_found],
                ["Exceptions", self.exception_files],
                ["# Sheets", self.num_sheets],
                ["Empty", self.empty_sheets],
                ["Date Issues", self.date_format_issue],
                ["Heading", self.headings_found],
                ["No Heading", self.headings_not_found]]

        output = tabulate(table, tablefmt="plain")
        return output


info = Info()

for file in os.listdir(RESULT_DIRECTORY):

    info.num_files += 1

    path = RESULT_DIRECTORY + file

    file_type = pathlib.Path(path).suffix

    if file_type in UNSUPPORTED_FILE_TYPES:
        info.unsupported_files.append(file)
        continue

    file_hash = sha256sum(path)
    if file_hash in info.file_hashes:
        info.duplicate_files_found += 1
        continue
    info.file_hashes.add(file_hash)

    try:
        file = ExcelFile(path)
    except Exception as e:
        info.exception_files += 1
        continue

    file.build_html()
    info.supported_file.append(file)

print(info)

unsorted_list = info.winners.values()
sorted_list = sorted(unsorted_list, key=lambda runner: runner.wins)
sorted_list.reverse()

table_list = []
for runner in sorted_list:
    table_list.append(runner.convert_to_list())

table = tabulate(table_list, tablefmt="plain")
print(table)

#remove double spaces, remove wheelchair guy jim corbet


# add the names as key with first name and surname in alphabetical order that way there should be no duplicates...