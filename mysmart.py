"""
This module will log into Smartsheet application and displays the current project status for each NEBS project.
Data will be written to an Excel spreadsheet located under the working directory's Results folder.

@author Kenneth Chen
"""


import datetime
import os
import re
import smartsheet
import logging
import pandas as pd


# This will only log the message for this module.  It prevents the 3rd party module log messages from appearing.
logger = logging.getLogger(__name__)
logger.setLevel(level=logging.DEBUG)

handler = logging.StreamHandler()
handler.setLevel(logging.DEBUG)
logger.addHandler(handler)

# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"
ACCESS_TOKEN = "30c7rnps712634rxa388xu09e0"

# The status sheet ID where all the NEBS projects information are stored.  This is maintained by Brent Taira.
NEBS_STATUS_SHEET_ID = 3151871180334980

# The status sheet ID where Kenneth Chen keeps track of active projects.
# MY_NEBS_STATUS_SHEET_ID = 7474637383722884
MY_TEST_SHEET_ID = 2363014939731844

# Workspace IDs

NEBS_WORKSPACE_IDS = [4792839738550148,   # NEBS Project Status
                      5745934320592772,   # NEBS Projects (eRAT)
                      382139716921220]    # NEBS Projects (Targa)

# NEBS_WORKSPACE_IDS = [4792839738550148]  # NEBS Project Status
# NEBS_WORKSPACE_IDS = [5745934320592772]  # NEBS Projects (eRAT)
# NEBS_WORKSPACE_IDS = [382139716921220]   # NEBS Projects (Targa)
# NEBS_WORKSPACE_IDS = [1043569512343428]  # Archived Projects
SG_WORKSPACE_IDS = [3517256463345540]     # IoT Project Status


def get_workspaces(ss_client):
    """
    Returns a list of Workspace objects

    :param ss_client: client
    :return: List of Workspace objects
    :rtype: Workspace
    """
    return ss_client.Workspaces.list_workspaces(include_all=True)  # optional argument page_size = 100, page = 1


# Returns a workspace object
def get_workspace_by_id(ss_client, w_id):
    """
    Returns the Workspace object given its id.

    :param ss_client: Access Token
    :param w_id:
    :return:
    """
    # Returns a workspace object with all sheets information populated
    return ss_client.Workspaces.get_workspace(w_id, load_all=True, include=["ownerInfo", "source"])


# Returns an array of sheet ids
def get_sheets_from_workspace(ws):
    arr = []  # temporary array of sheet ids
    ws_sheets = ws.sheets  # Array of Sheet objects

    if len(ws.sheets) > 0:
        for ws_sheet in ws.sheets:
            arr.append(ws_sheet.id)
    return arr


# Returns a Sheet object given a Sheet id
def get_sheet_by_id(ss, s_id):
    """
    Returns a Sheet object given a Sheet id

    :param ss:
    :param s_id:
    :return: Sheet object for that id
    :rtype: Sheet
    """
    return ss.Sheets.get_sheet(s_id, page_size=1000, include=["discussions",
                                                              "attachments",
                                                              "format",
                                                              "filters",
                                                              "ownerInfo",
                                                              "source",
                                                              "rowIds",
                                                              "rowNumbers",
                                                              "columnIds",
                                                             ])


# Returns all Sheets accessible by user
def get_all_sheets(ss):
    """
    Lists all the Sheets accessible by user.

    :param ss: Access Token
    :return:
    """
    response = ss.Sheets.list_sheets(include_all=True)
    return response.data


# Display all the parameters in the Sheet object
def show_sheet_parameters(ss, sh):
    print("show_sheet_param: \n {}".format(sh))
    print("Sheet ID = {}, Sheet Name = {}, Sheet Version = {}".format(sh.id, sh.name, sh.version))
    print("totalRowCount = {}".format(sh.totalRowCount))
    print("accessLevel = {}".format(sh.totalRowCount, sh.accessLevel))
    print("projectSettings = {}".format(sh.projectSettings))
    print("effectiveAttachmentOptions = {}".format(sh.effectiveAttachmentOptions))
    print("readOnly = {}".format(sh.readOnly))
    print("ganttEnabled = {}".format(sh.ganttEnabled))
    """
          "dependenciesEnabled = {}"
          "resourceManagementEnabled = {}"
          "favorite = {}"
          "showParentRowsForFilters = {}"
          "userSettings = {}"
          "ownerId = {}, owner = {}"
          "permalink = {}"
          "source = {}"
          "createdAt = {}, modifiedAt = {}"
          "columns = {}"
          "rows = {}"
          "discussions = {}"
          "attachments = {}"
          "fromId = {}"
    """


# Returns the workspace id
def get_workspace_id(workspace):
    logger.debug("get_workspace_id: type(workspace.id) = {}, workspace.id = {}".format(type(workspace.id), workspace.id))
    return workspace.id


# Return an array of workspace id
def get_workspaces_id(workspaces):
    w_arr = []
    for w in workspaces:
        logger.debug("get_workspaces_id: w.id = {}".format(w.id))
        w_arr.append(get_workspace_id(w))
    return w_arr


# Displays all the attributes of each workspace
def show_workspaces(sheet):
    response = sheet.Workspaces.list_workspaces(include_all=True)  # optional argument page_size = 100, page = 1
    ws = response.data
    for w in ws:
        logger.debug("show_workspaces: ID {}, name {}".format(w.id, w.name))


# Returns the eRAT number.
#
# Input: String
def get_erat_number(erat_str):
    if len(erat_str) > 0:
        pattern = re.compile(r'e(\d+):')
        if pattern.search(erat_str):
            result = pattern.match(erat_str).group(0)
            # logger.debug("get_erat_number: {}".format(result[1:5]))
            return result[1:5]
    return erat_str


# Returns the TAP number.
# In the ERAT format, it will be e1234:
# In the TARGA format, it will be T-5678:
#
# Input: String
def get_tap_number(tap_str):
    result = tap_str
    if len(tap_str) > 0:
        erat_pattern = re.compile(r'e(\d+)\D*:')
        targa_pattern = re.compile(r'T-(\d+):')
        if erat_pattern.search(tap_str):
            result = erat_pattern.match(tap_str).group(1)
        elif targa_pattern.search(tap_str):
            result = targa_pattern.match(tap_str).group(1).lstrip("0")
    # logger.debug("get_tap_number: result = {}".format(result))
    return result


# Returns the date
#
# Input: String in format of yyyy-mm-ddThh:mm:ss
# Output: datetime object
def get_date_obj(str):
    if re.compile('(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})').search(str):
        return datetime.datetime.strptime(str, '%Y-%m-%dT%H:%M:%S')
    elif re.compile('(\d{4}-\d{2}-\d{2})').search(str):
        return datetime.datetime.strptime(str, '%Y-%m-%d')


# Returns the date in the format: mm/dd/yyyy
#
# Input: date in string format (yyyy-mm-dd)
# Output: date in string format (mm/dd/yyyy)
def normalize_date(str):
    result = str

    # Looking for a yyyy-mm-dd pattern (e.g. 2017-09-14)
    if re.compile('(\d{4}-\d{2}-\d{2})').search(str):
        result = datetime.datetime.strptime(str, '%Y-%m-%d').strftime('%m/%d/%Y')
    return result


def normalize_cell(cell):
    result = ""
    if cell is not None:
        result = cell.display_value
        # logger.debug("normalize_cell: result = {}".format(result))
    return result


def normalize_tap_number(tap_str):
    """
    Returns the string with all the letters and leading zeros stripped.
    For example,
        Input: Targa-123
        Output: 123

    :param str tap_str: String with letters, numbers, and possibly symbols
    :return: The test plan number
    :rtype: string
    """
    result = tap_str

    # logger.debug("normalize_tap_number: tap_str = {}, type(tap_str) = {}".format(tap_str, type(tap_str)))
    if tap_str is not None:
        # Convert number types to str
        tap_str = str(tap_str)
        # Looking for a Targa-123 pattern.
        pattern = re.compile(r'Targa-(\d+)')
        if pattern.search(tap_str):
            result = pattern.match(tap_str).group(1)
        # logger.debug("normalize_tap_number: result = {}".format(result))
    return result


def erat_status(erat_num, ref_sheet, column_map, category):
    """
    Return an array of the following column values:
    Priority
    TAP number
    Category (NEBS)
    Project Name
    Project Manager

    :param ss_client:
    :param sheet:
    :param category:
    :return: A list containing Priority, TAP Number, Category, Project Name, and Project Manager.
    :rtype: list
    """
    # logger.debug("erat_status: erat_num = {}".format(erat_num))
    if re.compile('[0-9]').search(erat_num):
        for row in ref_sheet.rows:

            tap_number_cell = get_cell_by_column_name(column_map, row, "ERAT#")
            tap_number = normalize_tap_number(tap_number_cell.value)

            if erat_num == tap_number:

                priority_cell = get_cell_by_column_name(column_map, row, "Priority")
                priority = priority_cell.display_value  # value variable displays as a float

                name_cell = get_cell_by_column_name(column_map, row, "Project Name (link to status)")
                name = name_cell.value

                pm_cell = get_cell_by_column_name(column_map, row, "NEBS PM")
                pm = pm_cell.display_value  # value variable displays as an e-mail address.

                return [priority, erat_num, category, name_cell, pm]

        # ERAT number does not match anything in the NEBS Status sheet
        # Data returned with some entries blank.
        # return ["", erat_num, category, "", "", ""]
    else:
        return None


def sg_status(ss_client, sheet, category):
    """
    Return an array of the following column values:
    Priority
    TAP number
    Category (SG)
    Project Name
    Project Manager

    :param Smartsheet ss_client: base client object
    :param Sheet sheet: Specific sheet
    :param str category: NEBS or SG
    :return: A list containing Priority, TAP Number, Category, Project Name, and Project Manager.
    :rtype: list
    """
    sheet_column_map = build_column_map(ss_client, sheet)

    # Get the Sheet's owner user id
    user_id = sheet.owner_id
    user_obj = ss_client.Users.get_user(user_id)
    user_name = user_obj.first_name + " " + user_obj.last_name
    # logger.debug("sg_status: user_name = {}".format(user_name))
    project_code_name = normalize_cell(get_cell_from_col_row(sheet_column_map, sheet, "Standard Section", 1))
    project_id = normalize_cell(get_cell_from_col_row(sheet_column_map, sheet, "Standard Section", 2))
    project_name = project_code_name + " " + project_id
    return ['', '', category, project_name, user_name]


def get_cell_by_column_name(column_map, row, column_name):
    """
    Helper function to find the cell in a row

    :param dict column_map: Dictionary with column name as key and column id as value.
    :param Row row: Row where the cell is located.
    :param str column_name: Column where the cell is located.
    :return: Cell object at that specific column and row.
    :rtype Cell:
    :except KeyError: If the key is not found in the dictionary.
    """
    try:
        column_id = column_map[column_name]
        return row.get_column(column_id)
    except KeyError:
        return None


def build_column_map(ss_client, sheet):
    """
    The API identifies columns by Id, but it's more convenient to refer to column names.

    :param Smartsheet ss_client: base client object
    :param Sheet sheet: Specific sheet
    :return col_map: Dictionary with column name as key and column id as value
    :rtype dict:
    """
    col_map = {}
    # Iterate through all columns and stores the column ID
    for col in sheet.columns:
        col_map[col.title] = col.id
    # logger.debug("build_column_map: col_map = {}".format(col_map))
    return col_map


# Returns a test date given an user provided function
# Examples of functions could be min() or max()
def test_date(ss_client, sheet, func):
    date = ""
    date_arr = []

    # Get all the column titles from this sheet
    proj_column_map = build_column_map(ss_client, sheet)

    for row in sheet.rows:
        start_date_cell = get_cell_by_column_name(proj_column_map, row, "Start")
        if start_date_cell is not None:
            start_date = start_date_cell.value
            date_arr = str_to_date(start_date, date_arr)

        finish_date_cell = get_cell_by_column_name(proj_column_map, row, "Finish")
        if finish_date_cell is not None:
            finish_date = finish_date_cell.value
            date_arr = str_to_date(finish_date, date_arr)

    # Find the latest date in the array which is closest to today.
    if len(date_arr) > 0:
        # date_obj = max(dt for dt in date_arr)
        date_obj = eval(func)(dt for dt in date_arr)
        date = date_obj.strftime("%m/%d/%Y")
        # logger.debug("get_test_date: date = {}".format(date))
    return date


# Looks in the Start date column and returns the earliest date
def first_test_date(ss_client, sheet):
    return test_date(ss_client, sheet, "min")


# Looks in the Finish date column and returns the latest date
def last_test_date(ss_client, sheet):
    return test_date(ss_client, sheet, "max")


# Convert a date in string format to a datetime object.
def str_to_date(date_str, date_arr):
    if date_str is None or date_str == '':
        date_str = ""
    else:
        if re.compile('[0-9]').search(date_str):
            date_obj = get_date_obj(date_str)
            date_arr.append(date_obj)
    return date_arr


def get_cell_from_col_row(column_map, sheet, column_name, row_number):
    """
    Get the Cell object given a column name and specific row number

    :param map column_map: All the column
    :param Sheet sheet: Specific sheet
    :param str column_name: Column name where the cell is located
    :param int row_number: Row number where the cell is located (min 1)
    :return: the Cell object at the specified column and row
    :rtype: Cell
    """
    for row in sheet.rows:
        # logger.debug("get_cell_from_col_row: row.row_number = {}, row_number = {}".format(row.row_number, row_number))
        if row.row_number == row_number:
            return get_cell_by_column_name(column_map, row, column_name)


# Get the completion percentage from a specific cell on the sheet.
def completion(ss, sheet, column_name, row_number):
    result = ""
    proj_column_map = build_column_map(ss, sheet)

    for row in sheet.rows:
        if row.row_number == row_number:
            complete_cell = get_cell_by_column_name(proj_column_map, row, column_name)
            if complete_cell is not None:
                result = complete_cell.display_value  # value variable displays as a float (e.g. 88% is 0.88)
                if result is None:
                    result = "0%"
    return result


def get_excel_header():
    return ["Priority",
            "ERAT/TARGA",
            "NEBS/SG",
            "Project Name",
            "Completion",
            "Project Manager",
            "Start Date",
            "Last Test Date"]


# Update the Smartsheet with the contents of the Dataframe.
#
# @input ss: Smartsheet client where ACCESS_TOKEN has been instantiated.
# @input id: Smartsheet Sheet ID
# @input df: Dataframe object
def update_smartsheet(ss, id, df):

    # Load the sheet
    sheet = get_sheet_by_id(ss, MY_TEST_SHEET_ID)
    # sheet = ss.Sheets.get_sheet(MY_TEST_SHEET_ID, page_size=1000)
    # rows = sheet.rowIds()
    row_id = 0
    row_index = 0
    column_map = build_column_map(ss, sheet)
    logger.debug("update_smartsheet: column_map = {}".format(column_map))

    for df_row in df.itertuples(index=False):
        logger.debug("update_smartsheet: {}".format(df_row))

        priority = df_row[0]
        tap_num = df_row[1]
        category = df_row[2]
        project_name = df_row[3]
        complete = df_row[4]
        pm = df_row[5]
        start_date = df_row[6]
        last_test_date = df_row[7]
        logger.debug("update_smartsheet: [{},{},{},{},{},{},{},{}]".format(
                        priority,
                        tap_num,
                        category,
                        project_name,
                        complete,
                        pm,
                        start_date,
                        last_test_date))

        # Build the Smartsheet Cell
        priority_cell = build_cell(ss, column_map['Priority'], priority)
        tap_num_cell = build_cell(ss, column_map['ERAT'], tap_num)
        category_cell = build_cell(ss, column_map['NEBS/SG'], category)
        project_name_cell = build_cell(ss, column_map['Project Name'], project_name)
        complete_cell = build_cell(ss, column_map['Completion'], complete)
        pm_cell = build_cell(ss, column_map['Project Manager'], pm)
        start_date_cell = build_cell(ss, column_map['Start Date'], start_date)
        last_test_date_cell = build_cell(ss, column_map['Last Test Date'], last_test_date)

        arr_cells = [priority_cell,
                     tap_num_cell,
                     category_cell,
                     project_name_cell,
                     complete_cell,
                     pm_cell,
                     start_date_cell,
                     last_test_date_cell]

        # Build the Smartsheet Row
        # row_id = rows[row_index]
        new_row = build_row(ss, row_id, arr_cells)


# Create a new Smartsheet cell
def build_cell(ss, column_id, value, strict=True):
    new_cell = ss.models.Cell()
    new_cell.column_id = column_id
    new_cell.value = value
    new_cell.strict = strict
    return new_cell


# Create a new Smartsheet row
def build_row(ss, row_id, arr_cells):
    new_row = ss.models.Row()
    new_row.id = row_id
    for cell in arr_cells:
        new_row.cells.append(cell)
    return new_row


# Auto-generate a filename based on the current date and time.
#
# Returns the directory and filename as a string.
def generate_filename(str="", extension = '.xlsx'):
    current_date_time = datetime.datetime.now()
    strformat = str + "_" + current_date_time.strftime("%Y%m%d_%H%M%S")
    dir = os.path.dirname(os.path.abspath(__file__)) + os.sep + "Results" + os.sep
    str_path = dir + strformat + extension
    return str_path


# Returns a dataframe of project status when the workspace ID(s) are provided.
# Input: Workspace ID(s) Array of Integers
# Output: data_set Dataframe
def generate_dataframe_from_workspace(ss_client, workspace_ids, data_set,
                                      category = "NEBS",
                                      ref_sheet = None,
                                      ref_column_map = None):

    results_data = None

    for wk_id in workspace_ids:
        ws = get_workspace_by_id(ss_client, wk_id)
        arr_sheet_id = get_sheets_from_workspace(ws)
        logger.debug("generate_dataframe_from_workspace: arr_sheet = {}".format(arr_sheet_id))

        for i_sheet_id in arr_sheet_id:
            sheet = get_sheet_by_id(ss_client, i_sheet_id)
            logger.debug("generate_dataframe_from_workspace: sheet.id = {}, sheet.name = {}".format(sheet.id, sheet.name))
            if category == "NEBS":
                tap_num = get_tap_number(sheet.name)

            if ref_sheet is not None:
                """
                Using the ERAT number, retrieve data from NEBS master sheet
                results_data is list containing [Priority, 
                                                 ERAT#, 
                                                 Project Name (link to status)
                                                 NEBS PM]
                """
                if category == "NEBS":
                    results_data = erat_status(tap_num, ref_sheet, ref_column_map, category)
            elif category == "SG":
                results_data = sg_status(ss_client, sheet, category)

            if results_data is not None:
                # Skips any sheet that does not contain numbers.
                # The purpose is to filter out the Status Template sheet.
                if category == "NEBS":
                    if re.compile('[0-9]').search(tap_num):
                        # Retrieve data from inside the sheet
                        first_date = first_test_date(ss_client, sheet)
                        last_date = last_test_date(ss_client, sheet)
                        complete = completion(ss_client, sheet, "Standard Section No.", 2)
                elif category == "SG":
                    complete = completion(ss_client, sheet, "Standard Section", 4)
                    first_date = first_test_date(ss_client, sheet)
                    last_date = last_test_date(ss_client, sheet)

                results_data.insert(4, complete)
                results_data.append(first_date)
                results_data.append(last_date)
                logger.debug(results_data)

                data_set.append(results_data)  # for Dataframe
    return data_set


# Generates an excel spreadsheet and saves the dataframe.
# Input: Dataframe
# Output: None
def generate_excel(df, category):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(generate_filename(category), engine='xlsxwriter')

    if category == 'NEBS':
        df['Project Name'] = df['Project Name'].apply(lambda x: x.value)  # Retrieve string value from the Cell object

        # Sort the Dataframe by category (NEBS/SG) then ERAT number
        df_sort = df.sort_values(axis=0, by=['NEBS/SG', 'ERAT/TARGA'])
        logger.debug(df_sort.to_string())
        df = df_sort

    # Write Dataframe to Excel
    df.to_excel(writer, index=False)


# Retrieves the project status for NEBS and writes the data into an Excel spreadsheet.
def nebs():
    # Initialize variables
    data_set = []  # Used for Dataframe

    # Initialize client
    logger.info("Starting nebs()...")
    ss_client = smartsheet.Smartsheet(ACCESS_TOKEN)

    # Load the entire nebs master sheet
    nebs_master_sheet = ss_client.Sheets.get_sheet(NEBS_STATUS_SHEET_ID, page_size=1000)
    ref_column_map = build_column_map(ss_client, nebs_master_sheet)

    data_set = generate_dataframe_from_workspace(ss_client, NEBS_WORKSPACE_IDS, data_set, "NEBS", nebs_master_sheet, ref_column_map)

    # Create Dataframe using the data_set and column headers
    results_head = get_excel_header()  # Get the excel column labels
    df = pd.DataFrame.from_records(data_set, columns=results_head)
    generate_excel(df, "NEBS")
    """
    # Write Dataframe to Smartsheet
    # update_smartsheet(ss, MY_TEST_SHEET_ID, df_sort)

    # Get the xlsxwriter objects from the dataframe writer object.
    # workbook = writer.book
    # worksheet = writer.sheets['Sheet1']
    """


def smartgrid():
    # Initialize variables
    data_set = []  # Used for Dataframe

    # Initialize client
    logger.info("Starting smartgrid()...")
    ss_client = smartsheet.Smartsheet(ACCESS_TOKEN)
    data_set = generate_dataframe_from_workspace(ss_client, SG_WORKSPACE_IDS, data_set, category="SG")

    # Create Dataframe using the data_set and column headers
    results_head = get_excel_header()  # Get the excel column labels
    df = pd.DataFrame.from_records(data_set, columns=results_head)
    generate_excel(df, "SG")


def test():
    # Initialize variables
    # s_id = 2681468645336964  # Sheet ID for ERAT 6373 Tomahawk
    # s_id = 2868915110995844  # Sheet ID for ERAT 6759.  Contains date field = "n/a"
    # s_id = 3183397716682628  # Sheet ID for ERAT 6671.
    # s_id = 4294598903261060  # Sheet ID for Status Template (empty sheet)
    # s_id = 1751652114950020  # Sheet ID for T-0232 Firepower 7010
    # s_id = 6008318771652484  # Sheet ID for ERAT 6257a NCS-F
    # s_id = MY_TEST_SHEET_ID    # Sheet ID for Test
    s_id = 3432304224823172    # Sheet ID for Coronado Witness SG2 Test Plan and Status

    # Initialize client
    logger.info("Starting test() by instantiating the Smartsheet client using ACCESS_TOKEN.")
    ss_client = smartsheet.Smartsheet(ACCESS_TOKEN)
    sheet = get_sheet_by_id(ss_client, s_id)
    logger.debug("test: s.id = {}, s.name = {}".format(sheet.id, sheet.name))
    sg_status(ss_client, sheet, "SG")



def main():
    nebs()
    # smartgrid()

if __name__ == '__main__':
    while (True):
        choice = input("Select Menu:\n"
                       "(M)ain\n"
                       "(N)EBS\n"
                       "(S)martGrid\n"
                       "(T)est\n"
                       "E(x)it\n")
        if choice.lower() == 't':
            test()
        elif choice.lower() == 'm':
            main()
        elif choice.lower() == 'n':
            nebs()
        elif choice.lower() == 's':
            smartgrid()
        elif choice.lower() == 'x':
            exit()
        else:
            print("Invalid menu selection.")




