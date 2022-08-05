from auth import spreadsheet_service
from auth import drive_service

"""The auth workflow contains the path to credentials as well as the
    initial build of the spreadsheet and drive services"""

"""First off I am going to create a sheet class that will hold the various
funtions I want to use on the regs"""


"""TODO: So far using this as an import is really nice. I still need to add a few more methods
and explore some more of the drive api I am going to use this docstring to hold method ideas
1. Applying formatting, 2. applying to discontigious ranges, 3. reading discontigious ranges 4. New tabs"""

class Sheet:

    # Not really sure what needs to go here
    def __init__(self):
        pass

    # This creates a new sheet
    @staticmethod
    def create(title, shared_with=None):
        """TODO: I need to come up with a better way to push permissions
        I am also having difficulty deciding how to utilize the spreddy_id var
        in other functions of the class"""
        sheet = spreadsheet_service.spreadsheets()
        spreadsheet_properties = {
            'properties': {
                'title': title
            }
        }
        spreadsheet = sheet.create(body=spreadsheet_properties,
                                   fields='spreadsheetId').execute()
        spreddy_id = spreadsheet.get('spreadsheetId')
        permission1 = {
            'type': 'user',
            'role': 'writer',
            'emailAddress': shared_with
        }
        drive_service.permissions().create(fileId=spreddy_id, body=permission1).execute()
        return spreddy_id

    @staticmethod
    def read_sheet_contigious(sheet_key, sheet_name='Sheet1', sheet_range=None):
        if sheet_range:
            print("Sheet_range if executed")
            sheet = spreadsheet_service.spreadsheets()
            result = sheet.values().get(spreadsheetId=sheet_key,
                                        range=sheet_range).execute()
            values = result.get('values', [])
            print(values)
        else:
            print("Sheet_range else executed")
            sheet = spreadsheet_service.spreadsheets()
            result = sheet.values().get(spreadsheetId=sheet_key,
                                        range=sheet_name).execute()
            real_range = result.get('range', [])
            real_result = sheet.values().get(spreadsheetId=sheet_key,
                                             range=real_range).execute()
            values = real_result.get('values', [])
            print(values)

    @staticmethod
    def write_sheet(sheet_key, sheet_range, sheet_data: list) -> str:
        sheet = spreadsheet_service.spreadsheets()
        values = [
                sheet_data
        ]
        body = {
            'values': values
        }
        result = sheet.values().update(
            spreadsheetId=sheet_key, range=sheet_range,
            valueInputOption='USER_ENTERED', body=body).execute()
        print('{0} Cell(s) updates.'.format(result.get('updatedCells')))

    @staticmethod
    def create_new_tab(sheet_key, tab_title=None):
        sheet = spreadsheet_service.spreadsheets()
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': tab_title
                    }
                }
            }
            ]
        }
        sheet.batchUpdate(spreadsheetId=sheet_key,
                          body=request_body).execute()


