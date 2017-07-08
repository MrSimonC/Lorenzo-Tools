import os
import custom_modules.csv_tools as csv_tools
from custom_modules.mssql import QueryDB
from custom_modules.xlsx import XlsxTools
from custom_modules.excel_com import Excel
__version__ = 0.2


class FileToDB:
    """
    Process any file (.xlsx, xlsb, xls, csv) into a MS SQL database using dbo.Import_CSV_Files
    Requires Stored Procedure dbo.Import_CSV_files to exist in the database
    """
    def __init__(self, server, database, username, password):
        self.db = QueryDB(server, database, username, password)

    def import_file(self, input_file_path, output_pipe_folder, table, sheet_name='Sheet1', delimiter='|',
                    header_row_cell_value=''):
        if not os.access(input_file_path, os.R_OK):
            raise FileNotFoundError('Can\'t acccess file: ' + input_file_path)
        file_path_no_ext, file_extension = os.path.splitext(input_file_path)
        file_out = ''
        if file_extension == '.xlsb' or file_extension == '.xls' or file_extension == '.xlxm':
            # Prepare xlsb files into xlsx so openpyxl can read them
            xl = Excel()
            xl.open(input_file_path, read_only=True)
            save_to = os.path.join(output_pipe_folder, os.path.basename(input_file_path) + '.xlsx')
            xl.save_as(save_to, True)
            xl.close()
            file_out = os.path.join(output_pipe_folder, os.path.basename(input_file_path) + '.csv')
            xlsx = XlsxTools()
            xlsx.xlsx_to_csv(save_to, file_out, sheet_name, delimiter, header_row_cell_value=header_row_cell_value)
        if file_extension == '.xlsx':
            file_out = os.path.join(output_pipe_folder, os.path.basename(file_path_no_ext) + '.csv')
            xlsx = XlsxTools()
            xlsx.xlsx_to_csv(input_file_path, file_out, sheet_name, delimiter, header_row_cell_value=header_row_cell_value)
        if file_extension == '.csv':
            file_out = os.path.join(output_pipe_folder, os.path.basename(input_file_path))
            csv_tools.csv_to_pipe(input_file_path, file_out, delimiter)
        if file_out:
            self._sql_import_table(file_out, table)

    def _sql_import_table(self, file_path, table_name):
        sql = 'EXEC dbo.Import_CSV_files \'' + file_path + '\', \'' + table_name + '\''
        self.db.exec_sql(sql, True)
