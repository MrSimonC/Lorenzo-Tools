import os
from custom_modules import csv_tools
from custom_modules.xlsx import XlsxTools
from collections import OrderedDict


def service_point_data_pull_all(folder, results_file_output):
    """
    Go through a folder of xlsx files, collating data from each, outputting a csv of all data
    :param folder: folder with xlsx files to process
    :param results_file_output: csv output path
    :return: csv file
    """
    exclude = [
        # 'PRS - Physiotherapy 20150826b.xlsx',
    ]
    xlsx = XlsxTools()
    results = []
    for filename in os.listdir(folder):
        print(filename)
        _, ext = os.path.splitext(filename)
        if ext == '.xlsx' and '~' not in filename and filename not in exclude:
            try:
                file_dict = xlsx.dict_reader(os.path.join(folder, filename), 'clinics')
            except KeyError:
                try:
                    file_dict = xlsx.dict_reader(os.path.join(folder, filename), 'Clinics')
                except KeyError:
                    try:
                        file_dict = xlsx.dict_reader(os.path.join(folder, filename), 'Sheet1')
                    except:
                        raise
            for row in file_dict:
                results.append(OrderedDict([
                    ('Filename', filename),
                    # Cerner details
                    ('Cerner Template Name', row['Cerner Template Name'] if 'Cerner Template Name' in row else ''),
                    ('Cerner Resource', row['Cerner Resource'] if 'Cerner Resource' in row else ''),
                    ('Cerner Location', row['Cerner Location'] if 'Cerner Location' in row else ''),
                    ('Appointment Type (Concat)', row['Appointment Type (Concat)'] if 'Appointment Type (Concat)' in row else ''),
                    # Clinic
                    ('Lorenzo Clinic Name', row['Lorenzo Clinic Name']),
                    ('Clinic ID', row['Main Identifier'] if 'Main Identifier' in row
                        else row['Main Identifier TF1']),
                    # Clinician
                    ('Clinician Main Identifier', row['Clinician MainIdentifier'] if 'Clinician MainIdentifier' in row
                        else row['Clinician Main Identifier'] if 'Clinician Main Identifier' in row
                        else ''),
                    ('Care Provider', row['Care Provider'] if 'Care Provider' in row else ''),
                    # Treatment Function of clinic
                    ('Treatment Function', row['Treatment Function'] if 'Treatment Function' in row else ''),
                    # Session
                    ('Session Name', row['Session Name']),
                    ('From Date', row['From Date']),
                    ('Frequency', row['Frequency']),
                    # Location
                    ('Location ID', row['Most used Location'] if 'Most used Location' in row and
                        row['Most used Location'] != ''
                        else row['Lorenzo Location'] if 'Lorenzo Location' in row
                        else ''),
                    ('Location Lorenzo Main ID', row['Lorenzo Main ID'] if 'Lorenzo Main ID' in row else ''),
                    # Session continued
                    ('C&B Service Identifier',
                        ', '.join(sorted(row['C&B Service Identifier (again)'].replace(' ', '').split(',')))
                            if 'C&B Service Identifier (again)' in row and row['C&B Service Identifier (again)'] != ''
                        else ', '.join(sorted(row['C&B Service Identifier'].replace(' ', '').split(',')))
                            if 'C&B Service Identifier' in row and row['C&B Service Identifier'] != ''
                        else ''),
                    ('Slot Start Time', row['Slot Start Time']),
                    ('Slot End Time', row['Slot End Time']),
                    ('Priority', row['Priority']),
                    ('Applicable Apt Types', row['Applicable Apt Types'])
                ]))
    csv_tools.dict_to_csv(results_file_output, results)


def run():
    # folder = r'I:\Lorenzo Implementation\Documentation\System Configuration\Go-Live Build\Clinic Build'
    folder = r'U:\auto_delete'
    results_file_output = r'C:\Users\nbf1707\Desktop\Collation of Service Point data for compare1.csv'
    service_point_data_pull_all(folder, results_file_output)

# run()