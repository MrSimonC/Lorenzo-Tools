from custom_modules.mssql import QueryDB
from custom_modules.xlsx import XlsxTools
from custom_modules import excel_com
from custom_modules import csv_tools
import os


class LorenzoNonDCSIFFs:
    """
    Module to process Lorenzo Non IFF DCS files into SQL and check
    """
    def __init__(self):
        server = 'SERVER'
        database = 'DATABASE'
        username = 'USERNAME'
        password = 'PASSWORD'
        self.db = QueryDB(server, database, username, password)

    def sql_import_table(self, file_path, table_name):
        self.db.exec_sql('EXEC process.Import_CSV_files \'' + file_path + '\', \'' + table_name + '\'', True)


class NoDiff(Exception):
    pass


def main(process_xlsb=False):
    if process_xlsb:
        # Prepare xlsb files into xlsx so openpyxl can read them
        e = excel_com.Excel()

        # File locations
        dcs_users_original_file = r'I:\Lorenzo Implementation\Documentation\System Configuration\Data Migration\Users\DCS_USR_v6_LRC2.7_RVJ_20150904.xlsb'
        dcs_users_save_to = r'I:\Lorenzo Implementation\Documentation\System Configuration\Data Migration\non-dcs iffs\DCS_Users.xlsx'
        dcs_app_original_file = r'I:\Lorenzo Implementation\Documentation\System Configuration\Data Migration\Access Planning\DCS_AP_v3_LRC2.7_RVJ_20150904.xlsb'
        dcs_app_save_to = r'I:\Lorenzo Implementation\Documentation\System Configuration\Data Migration\non-dcs iffs\DCS_APPs.xlsx'
        dcs_locations_original_file = r'I:\Lorenzo Implementation\Documentation\System Configuration\Data Migration\Locations\DCS_Locations_v1_LRC2.7_RVJ_20150902.xlsb'
        dcs_locations_save_to = r'I:\Lorenzo Implementation\Documentation\System Configuration\Data Migration\non-dcs iffs\DCS_Locations.xlsx'

        # Check files exist!
        if not os.access(dcs_users_original_file, os.R_OK):
            print('Can\'t find Users DCS: ' + dcs_users_original_file)
            exit()
        if not os.access(dcs_app_original_file, os.R_OK):
            print('Can\'t find Access Plan Profile DCS: ' + dcs_app_original_file)
            exit()
        if not os.access(dcs_locations_original_file, os.R_OK):
            print('Can\'t find Locations DCS: ' + dcs_locations_original_file)
            exit()

        print('Converting DCS APPs to xlsx')
        e.open(dcs_app_original_file, read_only=True)
        e.save_as(dcs_app_save_to, True)
        e.close()

        print('Converting DCS Users to xlsx')
        e.open(dcs_users_original_file, read_only=True)
        e.save_as(dcs_users_save_to, True)
        e.close()

        print('Converting DCS Locations to xlsx')
        e.open(dcs_locations_original_file, read_only=True)
        e.save_as(dcs_locations_save_to, True)
        e.close()

    original_files_folder = r'I:\Lorenzo Implementation\Documentation\System Configuration\Data Migration\non-dcs iffs'
    pipe_output_folder = r'C:\Cerner Audit\lorenzoDCS'
    files_to_process = [
        {'fileInName': os.path.join(original_files_folder, 'LRD_SERVEVENTSTATUS.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'LRD_SERVEVENTSTATUS.csv'), 'sheetName': 'Sheet1',  'table_name': 'LRD_SERVEVENTSTATUS'},
        {'fileInName': os.path.join(original_files_folder, 'LRD_SERVICELOCN_IPED.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'LRD_SERVICELOCN_IPED.csv'), 'sheetName': 'Sheet1',  'table_name': 'LRD_SERVICELOCN_IPED'},
        {'fileInName': os.path.join(original_files_folder, 'LRD_SERVICELOCN_OP.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'LRD_SERVICELOCN_OP.csv'), 'sheetName': 'Sheet1',  'table_name': 'LRD_SERVICELOCN_OP'},
        {'fileInName': os.path.join(original_files_folder, 'LRD_SERVPOINT_IPED.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'LRD_SERVPOINT_IPED.csv'), 'sheetName': 'Sheet1',  'table_name': 'LRD_SERVPOINT_IPED'},
        {'fileInName': os.path.join(original_files_folder, 'LRD_SERVPOINT_OP.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'LRD_SERVPOINT_OP.csv'), 'sheetName': 'Sheet1',  'table_name': 'LRD_SERVPOINT_OP'},
        {'fileInName': os.path.join(original_files_folder, 'LRD_SERVPROVDET.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'LRD_SERVPROVDET.csv'), 'sheetName': 'Sheet1',  'table_name': 'LRD_SERVPROVDET'},
        {'fileInName': os.path.join(original_files_folder, 'LRD_TEAMS.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'LRD_TEAMS.csv'), 'sheetName': 'Sheet1',  'table_name': 'LRD_TEAMS'},
        {'fileInName': os.path.join(original_files_folder, 'DCS_Users.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'DCS_Users.csv'), 'sheetName': 'Users',  'table_name': 'DCS_USERS', 'headerRowCellValue': 'Record Number'},
        {'fileInName': os.path.join(original_files_folder, 'DCS_APPs.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'DCS_APP_Outpatient.csv'), 'sheetName': 'Outpatient', 'table_name': 'DCS_APP_OUTPATIENT', 'headerRowCellValue': 'profilename'},
        {'fileInName': os.path.join(original_files_folder, 'DCS_APPs.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'DCS_APP_Inpatient.csv'), 'sheetName': 'ElectiveAdmission', 'table_name': 'DCS_APP_INPATIENT', 'headerRowCellValue': 'ProfileName'},
        {'fileInName': os.path.join(original_files_folder, 'DCS_Locations.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'DCS_Locations.csv'), 'sheetName': 'Locations', 'table_name': 'DCS_LOCATIONS', 'headerRowCellValue': 'Record Number'},
        {'fileInName': os.path.join(original_files_folder, 'LRD_APGROUPING.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'LRD_APGROUPING.csv'), 'sheetName': 'Sheet1', 'table_name': 'LRD_APGROUPING'},
        {'fileInName': os.path.join(original_files_folder, 'mapping.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'mapping.csv'), 'sheetName': 'Sheet1', 'table_name': 'mapping'},
        {'fileInName': os.path.join(original_files_folder, 'location_mapping.xlsx'), 'fileOutName': os.path.join(pipe_output_folder, 'location_mapping.csv'), 'sheetName': 'mapping', 'table_name': 'location_mapping'},
    ]
    xlsx = XlsxTools()
    lr = LorenzoNonDCSIFFs()
    for file in files_to_process:
        print("Loading: " + file['table_name'])
        xlsx.xlsx_to_csv(file['fileInName'],
                       file['fileOutName'],
                       file['sheetName'],
                       delimeter='|',
                       header_row_cell_value=file['headerRowCellValue'] if 'headerRowCellValue' in file else '')

        # TODO: change xlsx_to_csv to the below
        # data = xlsx.dict_reader(file['fileInName'],
        #                         file['sheetName'],
        #                         header_row_cell_value=file['headerRowCellValue'] if 'headerRowCellValue' in file else '')
        # csv_tools.dict_to_csv(file['fileOutName'], data, '|')

        lr.sql_import_table(os.path.join(pipe_output_folder, file['fileOutName']), file['table_name'])
    #results = lr.db.exec_sql('EXEC [dbo].[dq_check]')
    #print(results)

main(False)