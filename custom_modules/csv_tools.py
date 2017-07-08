import csv
import os
__version__ = 1.0


def csv_to_pipe(file_in, file_out, delimiter='|', trim=False):
    if not os.access(file_in, os.R_OK):
        raise OSError('Can\'t read input file')
    csv_to_read = open(file_in, 'r')
    reader = csv.reader(csv_to_read)

    csv_to_write = open(file_out, 'w', newline='')
    writer = csv.writer(csv_to_write, delimiter=delimiter)

    for row in reader:
        writer.writerow([item.strip() if trim else item for item in row])
    csv_to_read.close()
    csv_to_write.close()


def csv_to_pipe_folder(folder_in, folder_out, delimiter='|', append_to_filename='_pipe'):
    # sac - don't think i'm using this function anywhere...yet!
    files_to_process = [os.path.join(folder_in, file) for file in os.listdir(folder_in)
                        if os.path.isfile(os.path.join(folder_in, file))]
    for file_path in files_to_process:
        path, filename_with_ext = os.path.split(file_path)
        file, ext = os.path.splitext(filename_with_ext)
        if ext == '.csv':
            pipes_output = os.path.join(folder_out, file + append_to_filename + '.csv')
            csv_to_pipe(file_path, pipes_output, delimiter, True)


def dict_to_csv(filename_csv_output, list_of_dict, delimiter=','):
    """
    Write a csv file from [{header: data}, {header: data} ...]
    :param filename_csv_output: output path
    :param list_of_dict: dictionary to transform into csv
    :param delimiter: delimiter
    """
    with open(filename_csv_output, 'w', newline='') as csv_file:
        writer = csv.DictWriter(csv_file, list(list_of_dict[0].keys()), delimiter=delimiter)
        writer.writeheader()
        for row in list_of_dict:
            writer.writerow(row)
