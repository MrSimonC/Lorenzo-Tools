import os
import pyperclip
import xml.etree.ElementTree as ET
import datetime


def get_files(folder):
    return [os.path.join(folder, file) for file in os.listdir(folder) if os.path.isfile(os.path.join(folder, file))]


def traverse_dtm_xml(xml_file):
    """
    Analyse Lorenzo Letter DTM file and output used merge fields
    :param xml_file: dtm file
    :return: [filename, merge_field1, merge_field2, ...]
    """
    dtm_in_xml = ET.parse(xml_file).getroot()  # or dtm_in_xml = ET.fromstring(xml_string)
    merge_fields = [os.path.basename(xml_file)]
    for doc_object in dtm_in_xml:
        if doc_object.tag == 'DocTemplateCode':
            merge_fields.append(doc_object.text)
        if doc_object.tag == 'DocTemplateName':
            merge_fields.append(doc_object.text.strip())
        if doc_object.tag == 'DocTemplateType':
            merge_fields.append(doc_object.text)
        if doc_object.tag == 'HIMObjects':
            for him_object in doc_object:
                merge_field = ''
                for him_object_name in him_object:
                    if him_object_name.tag == 'HIMObjName':
                        merge_field = him_object_name.text
                    if him_object_name.tag == 'HIMObjAttribute':
                        for him_object_attribute in him_object_name:
                            if him_object_attribute.tag == 'HIMObjAttName':
                                merge_fields.append(merge_field + '.' + him_object_attribute.text.strip())
    return merge_fields


def most_recent_files(folder):
    """
    Returns most recently modified files (with same name if you remove_no 16 chars off end)
    :param folder: folder to process
    :return: [[filename, dateobject], [filename, dateobject], ...]
    """
    files = get_files(folder)
    files_with_mod_dates = [[os.path.basename(file),
                            datetime.datetime.fromtimestamp(os.path.getmtime(file))]  # modified date
                           for file in files]
    files_with_mod_dates_uniq = []
    for file, mod_date in files_with_mod_dates:
        skip = False
        for file_dup, mod_date_dup in files_with_mod_dates:
            if file[:-16] == file_dup[:-16] and mod_date < mod_date_dup:
                skip = True
        if not skip:
            files_with_mod_dates_uniq.append([file, mod_date])
    return files_with_mod_dates_uniq


def process_folder_for_all_merge_fields(folder):
    """
    Produce full output of all merge fields in DTM files found in "folder"
    :param folder: Folder containing Lorenzo DTM (correspondence) files
    :return: Copies results to a clipboard for pasting into Excel
    """
    file_to_process = [file for file, mod_date in most_recent_files(folder)]
    file_fields = [traverse_dtm_xml(os.path.join(folder, file)) for file in file_to_process]
    file_fields_string = ['\n'.join(letter[0] + '\t' +
                                    letter[1] + '\t' +
                                    letter[2] + '\t' +
                                    letter[3] + '\t' +
                                    fields for fields in letter[4:]) for letter in file_fields]
    file_fields_output = '\n'.join(file_fields_string)
    pyperclip.copy('File\tID\tName\tType\tMerge Field\n' + file_fields_output)

# folder = r'C:\Users\nbf1707\Desktop\test'
folder_to_process = r'C:\Users\nbf1707\Desktop\All letters extract 5501 (24Nov15)'
process_folder_for_all_merge_fields(folder_to_process)