from xlrd import open_workbook, xldate_as_tuple, xldate_as_datetime
from tabulate import tabulate
from datetime import time
import pathlib
import os
import hashlib

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

RESULT_DIRECTORY = "output/"
UNSUPPORTED_FILE_TYPES = [".pdf", ".doc", ".docx"]


def strip_excessive_spaces(input_string):
    input_string = input_string.lstrip().rstrip()
    input_string = " ".join(input_string.split())
    return input_string


def sha256sum(filename):
    h = hashlib.sha256()
    b = bytearray(128*1024)
    mv = memoryview(b)
    with open(filename, 'rb', buffering=0) as f:
        for n in iter(lambda : f.readinto(mv), 0):
            h.update(mv[:n])
    return h.hexdigest()


class ExcelDateFormatException(Exception):
    pass


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
        self.handle = self.parent.workbook.sheet_by_index(index)
        self.index = index
        self.html_table = ""
        self.is_empty = self.is_empty_sheet()
        self.heading_row_found = False
        self.heading_row = -1
        self.winner = ""
        self.columns = {}

        if not self.is_empty:
            self.find_heading_row()
            if self.heading_row_found:
                self.winner = Runner(self.get_winner())

    def bootstrap_html_table(self):
        self.html_table = self.html_table.replace("<table>", "<table class=\"table table-striped table-hover table-sm\">")

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
                        time_raw = xldate_as_tuple(cell.value, self.parent.workbook.datemode)
                        value = str(time(*time_raw[3:]))
                    except Exception as e:
                        #info.date_format_issue += 1
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
                cell_value = self.get_cell_value(row, col)

                for heading in HEADER_STRINGS.values():
                    for pattern in heading:

                        if pattern == cell_value:
                            heading_matches_in_row += 1

                            if heading_matches_in_row > 2:
                                self.heading_row_found = True
                                self.heading_row = row
                                self.identify_columns()

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

    def get_cell_value(self, row, col):
        return str(self.handle.cell(row, col).value).lower().replace(" ", "")

    def identify_columns(self):

        for col in range(0, self.handle.ncols):

            cell_value = self.get_cell_value(self.heading_row, col)

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


class ExcelFile:

    def __init__(self, path):
        self.path = path
        self.html = ""
        self.sheets = []
        self.workbook = open_workbook(path)
        self.empty_sheets_count = 0

        # Only add non empty sheets
        for sheet_index in range(0, self.workbook.nsheets):
            sheet = Sheet(self, sheet_index)
            if sheet.is_empty:
                self.empty_sheets_count += 1
                continue
            else:
                self.sheets.append(sheet)

    def build_html(self):

        self.html = HTML_HEADER + CSS_SCRIPT

        for sheet in self.sheets:

            # Add a blank line between every sheet
            self.html += "<br>"

            try:
                self.html += sheet.build_html_table()
            except ExcelDateFormatException:
                return ""

        self.html += HTML_FOOTER

    def get_winners(self):
        winners = []
        for sheet in self.sheets:
            winners.append(sheet.winner)
        return winners


class Info:

    def __init__(self):
        self.files_count = 0

        self.file_hashes = set()
        self.headings_found = 0
        self.headings_not_found = 0
        self.sheets_count = 0
        self.empty_sheets_count = 0
        self.date_format_issue = 0

        self.supported_files_count = 0
        self.supported_files = []

        self.unsupported_files_count = 0
        self.unsupported_files = []

        self.duplicate_files_count = 0
        self.duplicate_files = []

        self.exception_files_count = 0
        self.exception_files = []
        self.winners = {}

    def add_winners(self, winners):

        for runner in winners:

            if runner is "":
                continue

            if runner.get_key() in self.winners.keys():
                self.winners[runner.get_key()].wins += 1
            else:
                runner.wins = 1
                self.winners[runner.get_key()] = runner

    def update_stats(self):

        self.files_count = len(self.unsupported_files) + len(self.supported_files)
        self.supported_files_count = len(self.supported_files)
        self.unsupported_files_count = len(self.unsupported_files)
        self.exception_files_count = len(self.exception_files)
        self.duplicate_files_count = len(self.duplicate_files)

        for file in self.supported_files:
            self.sheets_count += len(file.sheets)
            self.empty_sheets_count += file.empty_sheets_count
            self.add_winners(file.get_winners())

    def generate_winner_table(self):
        unsorted_list = info.winners.values()
        sorted_list = sorted(unsorted_list, key=lambda runner: runner.wins)
        sorted_list.reverse()

        table_list = []
        for runner in sorted_list:
            table_list.append(runner.convert_to_list())

        table = tabulate(table_list, tablefmt="plain")
        return table

    def __repr__(self):
        self.update_stats()

        table = [["# Files", self.files_count],
                ["Supported", self.supported_files_count],
                ["Unsupported", self.unsupported_files_count],
                ["Duplicates", self.duplicate_files_count],
                ["Exceptions", self.exception_files_count],
                ["# Sheets", self.sheets_count],
                ["Empty", self.empty_sheets_count],
                ["Date Issues", self.date_format_issue],
                ["Heading", self.headings_found],
                ["No Heading", self.headings_not_found]]

        output = tabulate(table, tablefmt="plain")
        return output


if __name__ == "__main__":

    info = Info()

    for file in os.listdir(RESULT_DIRECTORY):

        path = RESULT_DIRECTORY + file
        file_type = pathlib.Path(path).suffix

        if file_type in UNSUPPORTED_FILE_TYPES:
            info.unsupported_files.append(file)
            continue

        file_hash = sha256sum(path)
        if file_hash in info.file_hashes:
            info.duplicate_files.append(path)
            continue
        info.file_hashes.add(file_hash)

        try:
            file = ExcelFile(path)
            info.supported_files.append(file)
        except Exception as e:
            info.exception_files.append(path)
            continue

    print(info)
    print(info.generate_winner_table())
