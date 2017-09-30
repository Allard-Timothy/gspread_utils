"""Wrappers to the Google Spreadsheet API Gspread.
You can also find documentation on Gspread below:
http://gspread.readthedocs.org/en/latest/
"""

import gspread
from itertools import izip
from dteam import config
from oauth2client.client import SignedJwtAssertionCredentials

LOGGER = config.logging.getLogger(__name__)


def api(spreadsheet_name):
    """Provide an authorized instance of GoogleSpreadsheetWrapper.
    :param spreadsheet_name: name of spreadsheet to instantiate
    :return: Instance of spreadsheet class
    """
    return GoogleSpreadsheetWrapper(spreadsheet_name)


def _worksheet_api(worksheet):
    """Provide an instance of GoogleWorksheetWrapper
    :param worksheet: name of worksheet to instantiate
    :return: Instance of worksheet class
    """
    return GoogleWorksheetWrapper(worksheet)


def _auth():
    """Provide auth to gspread using OAuth2."""
    credentials = config.get().gspread
    scope = ['https://spreadsheets.google.com/feeds']
    cred = SignedJwtAssertionCredentials(scope=scope, **credentials)
    return gspread.authorize(cred)


def _get_col_letter(index):
    """Returns the alphanumeric column index for the given numerical column index.
    :param index: Integer representing column of spreadsheet
    :return: Aplhanumeric column
    >>> k = range(29)
    >>> _get_col_letter(len(k) - 1)
    'AC'
    """
    mapping = ('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
    base = 26
    if index < 0:
        raise ValueError('Index value must be greater than 0')
    quotient, remainder = divmod(index, base)
    if quotient > 0:
        return _get_col_letter(quotient - 1) + mapping[remainder]
    return mapping[remainder]


def _build_column_keys(values):
    """This takes in a list of values and uses the length of
    the list to determine the column range.
    :param values: A list of values used to determine end column letter
    :return: Aplhanumeric cell range
    >>> k = [x for x in range(29)]
    >>> _build_column_keys(k)
    'A1:AC1'
    """
    cell_range = 'A1:{col_letter}1'
    value = len(values) - 1
    cell_letter = _get_col_letter(value)
    return cell_range.format(col_letter=cell_letter)


def _build_value_range(values, current_row):
    """This takes in a list of values and  determines
    the cell range for row based on the length of values.
    :param values: A list of values used to determine cell range
    :param current_row: Empty row to be filled
    :return: Alphanumeric cell range used to write to row range
    """
    end_cell_letter = _get_col_letter(len(values) - 1)
    return 'A{current_row}:{end_col}{end_row}'.format(current_row=current_row,
                                                      end_col=end_cell_letter,
                                                      end_row=current_row)


def _build_col_range(values, column_letter):
    """Creates row range for specific column based on length of input.
    :param values: A list of values to determine cell range for column
    :param column_letter: Letter of column to be filled
    :return: Aplhanumeric cell range used to write to column range
    """
    end_row_num = len(values) + 1
    cell_list = '{col_letter}2:{col_letter}{row_num}'.format(col_letter=column_letter.upper(),
                                                             row_num=end_row_num)
    return cell_list


class GoogleSpreadsheetWrapper(object):

    """A Spreadsheet Wrapper to the Google Spreadsheet API Gspread."""

    def __init__(self, spreadsheet_name):
        self.client = _auth()
        self.spreadsheet = self.client.open(spreadsheet_name)

    def get_worksheet(self, sheet_name):
        """Returns specified worksheet from logged in Google account.
        :param sheet_name: Name of desired worksheet
        :return: Instance of the the worksheet
        """
        worksheet = self.spreadsheet.worksheet(sheet_name)
        return _worksheet_api(worksheet)

    def create_sheets(self, sheet_list):
        """Takes in list of worksheet names and creates worksheet for each
        element in sheet_list.
        :param sheet_list: Creates worksheets for spreadsheet derived from sheet_list
        """
        for k in location_list:
            self.spreadsheet.add_worksheet(k, rows=200, cols=200)


class GoogleWorksheetWrapper(object):

    """A Worksheet wrapper to the Google Spreadsheet API Gspread."""

    def __init__(self, worksheet):
        self.worksheet = worksheet

    def find_last_col(self, headers):
        """Finds the last populated column.
        :param headers: List of headers to determine last populated column
        :return: Letter of the last populated column
        """
        return _get_col_letter(len(headers) - 1)

    def find_last_row(self, col_value):
        """Finds the last populated row.
        :param col_value: Value of column to check last populated row
        :return: Last populated row for given column"""
        return len(self.worksheet.col_values(col_value))

    def dict_fill(self, data):
        """Will use dict keys and values to write to spreadsheet headers
        and the corresponding cell values. The range of cells to be written
        is determined by the last populated row in column A and the length
        of the data object passed in.
        :param data: Dict object where keys/values are header/cell values
        """
        last_row = self.find_last_row(1) + 1
        cell_range = _build_value_range(data, last_row)
        cells = self.worksheet.range(cell_range)
        for cell, value in zip(cells, values):
            cell.value = value
        self.worksheet.update_cells(cells)

    def get_all_records(self):
        """Return all records from worksheet in a list of dicts format.
        Each dict is a row containing column headers as keys matched
        with the cell values."""
        return self.worksheet.get_all_records(empty2zero=False)

    def get_all_values(self):
        """Return all values from worksheet in a list of lists format.
        Each sub-list corresponds to a row within the worksheet."""
        return self.worksheet.get_all_values()

    def list_fill(self, data, _range):
        """Writes a list of values to spreadsheet. The cells to be written
        are determined by last row in column A and then length of the data
        object passed in.
        :param data: List of values to be written
        """
        last_row = self.find_last_row(1) + 1
        cell_range = _build_value_range(data, last_row)
        cells = self.worksheet.range(cell_range)
        for cell, new_value in izip(cells, data):
            cell.value = new_value
        self.worksheet.update_cells(cells)

    def fill_headers(self, header_list):
        """The header_list parameter represents a list of values to be used
        as headers which will populate the first row of the worksheet.
        :param header_list: List of header values to be written to first row
        :return: Values written to first row of spreadsheet
        """
        header_range = _build_column_keys(header_list)
        self.list_fill(header_list, header_range)
        return header_list

    def build_sheet_by_keys(self, data, start_row, keys=None, headers=None):
        """Default behavior is to use the spreadsheets first row as the header
        keys to match with values. If provided, will use headers to fill in first
        row and match values accordingly, elements in headers must be present in dict as keys.
        :param data: List of dicts where keys are header values and values are cell values
        :param start_row: Integer of first row to be written
        :param keys: Optional list of keys to match against headers
        :param headers: Optional list of values to use as header values
        """
        headers = self.worksheet.row_values(1) if not headers \
            else self.fill_headers(headers)
        keys = set(headers) if not keys else set(keys)
        for row, item in enumerate(data, start=start_row):
            range_string = 'A{row}:{end_col}{row}'
            cell_letter = _get_col_letter(len(headers) - 1)
            cells = self.worksheet.range(range_string.format(row=row,
                                                             end_col=cell_letter))
            values = [item[key] for key in headers if key in keys]
            for cell, value in zip(cells, values):
                cell.value = value
            self.worksheet.update_cells(cells)

    def list_all_rows(self):
        """List all populated rows in the worksheet.
        :return: list of row urls pertaining to each row in the worksheet.
        """
        return self.worksheet.list_rows()

    def delete_row(self, row_url):
        """Delete the passed in row from the worksheet. The row_url
        will be provided by the list_all_rows function.

        :param row: row to be deleted
        """
        self.worksheet.delete_row(row_url)

    def delete_many_rows(self, row_url_list):
        """Delete all row_urls in the iterable row_url_list.

        :param row_url_list: list of row_urls returned from list_all_rows
        """
        for row_url in row_url_list:
            self.delete_row(row_url)
