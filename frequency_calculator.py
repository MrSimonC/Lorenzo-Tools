import csv
import datetime
from dateutil.relativedelta import relativedelta
import re


def adjust_date_to_frequency_letter(apply_date: datetime.datetime, frequency_letter: str):
    """
    Returns the future date from apply_date to match the frequency_letter
    e.g. 01/01/00 = monday, frequency_letter = t, return = 02/01/00
    :param apply_date: start date of the clinic
    :param frequency_letter: the letter of the week the first clinic should fall on
    :return: adjusted start date to match the day of the week letter
    """
    frequency_letter = frequency_letter.lower().strip()
    start_date_letter = day_into_letter(apply_date)
    for day_count in range(0, 7):
        if frequency_letter == start_date_letter:
            return apply_date
        apply_date += datetime.timedelta(days=1)
        start_date_letter = day_into_letter(apply_date)


def day_into_letter(date: datetime.datetime):
    """
    Takes a date, returns the NBT letter of the week
    :param date: date
    :return: m / t / w / h / f / s / u
    """
    return 'h' if date.strftime('%a') == 'Thu' else \
           'u' if date.strftime('%a') == 'Sun' else date.strftime('%a')[:1].lower()


def falls_on_holiday(start: datetime.datetime, holiday: datetime.datetime, frequency: int, week_or_month: str):
    week_or_month = week_or_month.lower().strip()
    if week_or_month != 'w' and week_or_month != 'm':
        raise ValueError
    if start > holiday:
        return False
    next_date = start
    while next_date <= holiday:
        # print(next_date.strftime('%d/%m/%Y'))
        if next_date == holiday:
            return True
        next_date = next_date + (datetime.timedelta(weeks=frequency) if week_or_month == 'w'
                                 else relativedelta(months=+frequency) if week_or_month == 'm'
                                 else datetime.timedelta(days=0))


def main():
    """
    Takes an input CSV file with the titles shown below, and outputs a csv with "TRUE/FALSE" in the
    'Falls on Target Date' column. Useful to find if clinics fall on a date e.g. bank holiday
    :return: csv file output
    """
    csv_input = r'C:\Users\nbf1707\Desktop\clinic input.csv'
    start_date_header = 'Session Start Date'
    frequency_header = 'Frequency'
    target_date_header = 'Target Date'
    true_false_header = 'Falls on Target Date'
    csv_output = r'C:\Users\nbf1707\Desktop\clinic output.csv'
    re_ew = r'(?:E)(\d+)([WM])'  # days = group 0, frequency = group 1, week or month = group 2
    fail_text = 'Can''t determine'

    csv_input_dict = csv.DictReader(open(csv_input))
    live_clinics = [x for x in csv_input_dict]
    for row in live_clinics:
        start_date = datetime.datetime.strptime(row[start_date_header], '%d/%m/%Y')
        frequency_raw = row[frequency_header]
        target_date = datetime.datetime.strptime(row[target_date_header], '%d/%m/%Y')

        days = frequency_raw.split(' ')[0] if frequency_raw.split(' ')[0] else fail_text
        frequency = int(re.search(re_ew, frequency_raw).group(1)) \
            if re.search(re_ew, frequency_raw) else fail_text
        w_or_m = re.search(re_ew, frequency_raw).group(2) \
            if re.search(re_ew, frequency_raw) else fail_text

        row[true_false_header] = 'False'
        for day in list(days):
            if falls_on_holiday(adjust_date_to_frequency_letter(start_date, day), target_date, frequency, w_or_m):
                row[true_false_header] = 'True'

    csv_object = csv.DictWriter(open(csv_output, 'w', newline=''), csv_input_dict.fieldnames + [true_false_header])
    csv_object.writeheader()
    csv_object.writerows(live_clinics)

if __name__ == '__main__':
    import time
    start_time = time.time()
    main()
    print(time.time() - start_time)