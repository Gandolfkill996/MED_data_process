# ----------------------------------------------------------------------
# 1. Imports
# ----------------------------------------------------------------------
import pandas as pd
from datetime import datetime
from collections import *
import os



# ----------------------------------------------------------------------
# 2. File path setup
# ----------------------------------------------------------------------
# The Excel file that contains all REACH cohort and visit data.
# NOTE: Change this path when moving between environments eg. address = "/Users/dentakugun/Downloads/REACH Visit Compliance.xlsx".
# Automatically detect Excel file in the same folder as this script
script_dir = os.path.dirname(os.path.abspath(__file__))
address = os.path.join(script_dir, "REACH Visit Compliance tl 2025-1013.xlsx")
print(address)

# ----------------------------------------------------------------------
# 3. Utility functions
# ----------------------------------------------------------------------
def read_xlsx(address,sheet):
    '''
    Read data excel sheet data
    :param address:
    :param sheet:
    :return:
    '''
    data = pd.read_excel(address,sheet_name=sheet)
    return data

def days_diff(Date1, Date2):
    '''
    calculate the difference between two date eg. Date1: 6/29/2015 Date2: 2/22/2019
    :param Date1:
    :param Date2:
    :return: Date2 - Date1
    '''
    # If one of them is NaT or NaN，return None
    if pd.isnull(Date1) or pd.isnull(Date2):
        return None

    # If date is string type，transfer to pandas Timestamp
    if isinstance(Date1, str):
        Date1 = pd.to_datetime(Date1, errors='coerce')
    if isinstance(Date2, str):
        Date2 = pd.to_datetime(Date2, errors='coerce')

    # if date is datetime.datetime，use directly
    if isinstance(Date1, (pd.Timestamp, datetime)) and isinstance(Date2, (pd.Timestamp, datetime)):
        return (Date2 - Date1).days

    # other situation will return None
    return None


def fix_date_format(date):
    '''
    Transfer date(read through pandas) into Timestamp format
    :param date:
    :return: format date
    '''
    if pd.isnull(date):
        date = pd.NaT
    else:
        date = pd.to_datetime(date)
    return date

# ----------------------------------------------------------------------
# 4. Retrieve cohort start/off-study information
# ----------------------------------------------------------------------
def get_visits_init_off_dates(address):
    """
    Combine initiation and off-study dates from both Original and New cohorts.
    Returns:
        visits_ini_off : dict
            {
                'ID': {
                    'init': <Timestamp>,
                    'off' : <Timestamp>,
                    'type': 'origin' or 'new'
                }
            }
    """
    # --- Read and clean Original cohort ---
    origin = read_xlsx(address,"4_Original Cohort")
    new = read_xlsx(address,"4_ New Cohort")

    # A dictionary include patients ID, initiate study date and off study date
    # Dictionary structure: {ID:{'init': initiate date; 'off': off study date}}
    # Some patients do not have off study date. Therefore, their off study date is dataframe.nan
    visits_ini_off = {}

    # Go through origin cohort sheet, record each patient's initiate and off date
    for i in range(len(origin)):
        ID, Site, Sc, Screening_days, Off_date, reason, init_date = origin.iloc[i]

        # In case, skip records without ID or initiate_date
        if pd.isnull(ID) or str(ID).strip() == "" or str(ID).lower() == "nan" or pd.isnull(init_date):
            continue

        # Transfer ID to prevent missmathch of number format
        if str(ID)[0] == '3':
            ID = '0'+str(ID)
        else:
            ID = str(ID)
        init_off_date={}

        # In original cohort, some patients who do not initiate study mark as 'Na - Sc' in HU init date
        if init_date == 'Na - Sc':
            continue
        # Fix date format and write into dictionary
        if ID not in visits_ini_off:
            init_off_date['init'] = fix_date_format(init_date)
            init_off_date['off'] = fix_date_format(Off_date)
            init_off_date['type'] = 'origin'
            visits_ini_off[ID] = init_off_date
        # If patient already exists in the dictionary, it means there are more than one patient cohort records
        else:
            print(f"ID: {ID} already have a new init date{init_date}, previous init date is: {visits_ini_off[ID]}")
            pass

    # Same as New cohort
    for i in range(len(new)):
        ID, Site, Consent, init_date, Off_date = new.iloc[i]
        if pd.isnull(ID) or str(ID).strip() == "" or str(ID).lower() == "nan":
            continue
        if str(ID)[0] == '3':
            ID = '0' + str(ID)
        else:
            ID = str(ID)
        init_off_date={}

        # In new cohort, some patients did not initiate study mark as N/A Off Study in HU initiation column, skip it
        if pd.isnull(init_date) or init_date=="N/A Off Study":
            # print(f"{ID} did not have initate visit!")
            continue
        if ID not in visits_ini_off:
            init_off_date['init'] = fix_date_format(init_date)
            init_off_date['off'] = fix_date_format(Off_date)
            init_off_date['type'] = 'new'
            visits_ini_off[ID] = init_off_date
        else:
            print(f"ID: {ID} already have a new init date{init_date}, previous init date is: {visits_ini_off[ID]}")
            pass
    return visits_ini_off


# ----------------------------------------------------------------------
# 5. Extract all visit records and compute relative days since initiation
# ----------------------------------------------------------------------
def get_vistis_intervals():
    '''
    Calculate each patient visit's date from intiate study date.
    visit_class: a dictionary which format is {ID1:{days_1, days_2 ...};ID2:{days_1, days_2 ...}}. days_x refers to how many days between initiate study date and visit date.
    No_regist_patient: prevent the situation that some visits' ID does not appear in Orignial and new cohort table.
    :return:
        visit_class : {ID: {"record": deque([days_since_init]), "type": cohort_type}}
        No_regist_patient : list of visit IDs without valid initiation date
    '''
    # address = "/Users/dentakugun/Library/CloudStorage/OneDrive-UniversityofCincinnati/Teresa/REACH Visit Compliance.xlsx"
    visits_record = read_xlsx(address,"5_Visit Dates")

    # Sort visits record first by ID and then by month
    visits_record = visits_record.sort_values(by=['record_id', 'month'])
    # Get visits_ini_off dictionary
    visits_ini_off = get_visits_init_off_dates(address)

    visit_class = {}
    No_regist_patient = []
    for i in range(len(visits_record)):
        record_id, site, redcap_event_name, visit_date,  month = visits_record.iloc[i]

        # In 5_Visit Dates table, some records' month missing unregular visit which is not included in calculation
        # Some records' month less than 0 means visit before initiate date which is not included in calculation
        # Some records' month is 0 refers initiate date. However, not all patients have initiate date record.
        # Therefore, use visits_ini_off from Orignial and new cohort table to get patient's initiate study date
        # Make sure ID has cohort info in 4_Original Cohort or 4_ New Cohort
        if record_id not in visits_ini_off:
            continue
        # Make sure all scheduled visit happens after initiate date
        if days_diff(visits_ini_off[record_id]['init'], visit_date) < 0:
            continue

        # Remove unschedule visit
        event = str(redcap_event_name).strip().lower() if pd.notnull(redcap_event_name) else ""
        allowed_prefixes = ("form_", "month_", "quarter_", "visit_month_","hydroxyurea_initat_arm_","month_0_hu_initiat_arm_")

        if not any(event.startswith(prefix) for prefix in allowed_prefixes):
            continue

        if record_id not in visits_ini_off.keys():
            No_regist_patient.append(record_id)
        else:
            if record_id not in visit_class:
                visit_class[record_id] = {}
                visit_class[record_id]['record']=deque([days_diff(visits_ini_off[record_id]['init'],visit_date)])
                visit_class[record_id]['type'] = visits_ini_off[record_id]['type']
            else:
                visit_class[record_id]['record'].append(days_diff(visits_ini_off[record_id]['init'],visit_date))

    return visit_class, No_regist_patient


def count_total_windows(study_days, type):
    '''
    calculate each individual maximum window number
    :param study_days: How many days of an individual from study initiate date to study off date
    :param type: 'Origin' or 'New'
    :return: total_windows for an ID
    '''
    # Aplly Cohort Specs rules in xlsx
    if type == 'origin':
        if study_days < 24 * 28 - 7:
            total_windows = (study_days + 7) // 28
        else:
            total_windows = 24 + ((study_days - (24 * 28 - 7)) + 14) // 28
    elif type == 'new':
        if study_days < 6 * 30 - 7:
            total_windows = (study_days + 7) // 30
        else:
            total_windows = 6 + ((study_days - (6 * 30 - 7)) + 14) // 30
    else:
        print(f"Unnormal type: {type}")
    return total_windows

# ----------------------------------------------------------------------
# 6. Main window-generation and compliance classification
# ----------------------------------------------------------------------
def calculation(current_date, cohort_type):
    """
    Build visit windows for each participant and mark visits as in/out-of-window.
    Includes COVID-19 pause handling and off-study truncation.

    Parameters
    ----------
    current_time : str
        Date string (e.g., '9/30/2025') used as analysis cutoff if off-study missing.
    cohort_type : str
        'origin' or 'new'

    :return: visit_count for each ID's each month's visit count according to 'new' or 'origin' cohort
    """
    # Get visits_ini_off dictionary
    visits_ini_off = get_visits_init_off_dates(address)
    # Get each patient visit's date from intiate study date.
    visit_class, No_regist_patient = get_vistis_intervals()
    # Keep covid start and end date in dictionary
    covid_interval ={"start":pd.to_datetime("3/1/2020"), "end":pd.to_datetime("1/1/2024")}

    # Keep each month's each individuals individual_visit_record in dictionary
    # The structure of visit_count is
    # {
    #   ID_1:{
    #           month_1:{"in_window":in_window, "out_window": out_window,"status": "status"},
    #           month_2:{"in_window":in_window, "out_window": out_window,"status": "status"}
    #           ...},
    #   ID_2:{
    #           month_1:{"in_window":in_window, "out_window": out_window,"status": "status"},
    #           month_2:{"in_window":in_window, "out_window": out_window,"status": "status"}
    #           ...},...
    #}
    visit_count={}
    for ID in visits_ini_off.keys():
        # If there is not visit record in 5_Visit Dates
        # or ID belongs to the type which we do not expect. Then skip
        if ID not in visit_class.keys() or visits_ini_off[ID]['type'] != cohort_type:
            continue

        # If there is no off study date for an individual, then set a current for their off study date
        if pd.isnull(visits_ini_off[ID]['off']):
            visits_ini_off[ID]['off'] = pd.to_datetime(current_date)

        # Calculate how many days between covid start date and ID's study initiate date, might be nefative
        covid_start_days = days_diff(visits_ini_off[ID]['init'],covid_interval['start'])
        # Calculate how many days between covid end date and ID's study off date, might be nefative
        covid_end_days = days_diff(visits_ini_off[ID]['init'],covid_interval['end'])
        # Calculate how many days between ID's study initiate and offf date
        study_days = days_diff(visits_ini_off[ID]['init'],visits_ini_off[ID]['off'])
        # calculate how many weeks(windows) for an individual
        total_weeks = count_total_windows(study_days, visits_ini_off[ID]['type'])
        # If there is no normal visit but initiate visit, force the total_weeks to 1
        total_weeks = max(1, total_weeks)
        # Keep each ID's every month(window) performance
        # individual_visit_record structure is
        # {
        #   month_1:{"in_window":in_window, "out_window": out_window,"status": "status"},
        #   month_2:{"in_window":in_window, "out_window": out_window,"status": "status"}
        #   ...},
        individual_visit_record = {}

        # If the wanted output id 'Origin'
        if visits_ini_off[ID]['type'] == 'origin':

            # Go through ID's each week(window)'s number
            for i in (range(total_weeks)):
                # initiate record
                # For 'Origin' ID, the boudary of windows' boundary differs at 24 month
                if i <=24:
                    individual_visit_record[i] = {"in_window": 0, "out_window": 0, "status": ''}
                    # calculate each window begin and end day
                    window_begin = (i)*28-7
                    window_end = (i)*31+7
                    # Determine if the window completely belongs to 'covid' period and mark windows
                    individual_visit_record[i]["status"] = 'norm'
                    if covid_start_days <= window_begin and window_end <= covid_end_days:
                        individual_visit_record[i]["status"] = 'covid'
                    # If current analyzing visit is out-window
                    while visit_class[ID]['record'] and visit_class[ID]['record'][0] < window_begin:
                        # Because visits before initiate date is not treat as regular visit,
                        # There is no out-window visit before month 0
                        if i > 0:
                            # Record the out-window visit between (n-1)th and (n)th window to (n-1)'s window
                            individual_visit_record[i-1]["out_window"] +=1

                        # remove out-window visit from the total visits record since is has alreay been counted
                        visit_class[ID]['record'].popleft()
                    # If current analyzing visit is in-window
                    while visit_class[ID]['record'] and window_begin <= visit_class[ID]['record'][0] <=window_end:
                        # Only record once if there are multiple in-window visit
                        if individual_visit_record[i]["in_window"] == 0:
                            individual_visit_record[i]["in_window"] += 1
                        # remove other in-window visits in current window
                        visit_class[ID]['record'].popleft()
                    # For month 0, all individual has visit as default
                    if i == 0:
                        individual_visit_record[i]["in_window"] = 1
                        continue
                # Same treatment for Original ID's month > 30.
                # The only difference is window's boundary calculation method
                elif i <=48:
                    individual_visit_record[i] = {"in_window": 0, "out_window": 0, "status": ''}
                    window_begin = (i)*28-14
                    window_end = (i)*31+14
                    individual_visit_record[i]["status"] = 'norm'
                    if covid_start_days <= window_begin and window_end <= covid_end_days:
                        individual_visit_record[i]["status"] = 'covid'
                    while visit_class[ID]['record'] and visit_class[ID]['record'][0] < window_begin:
                        individual_visit_record[i-1]["out_window"] =individual_visit_record[i-1]["out_window"]+1
                        visit_class[ID]['record'].popleft()
                    while visit_class[ID]['record'] and window_begin <= visit_class[ID]['record'][0] <=window_end:
                        if individual_visit_record[i]["in_window"] == 0:
                            individual_visit_record[i]["in_window"] += 1
                        visit_class[ID]['record'].popleft()

                # When month > 48, origin cohort ID requires to visit every 3 month
                else:
                    # If the month > 48, before the required visit month(51, 54...), skip these month. eg. 49, 50, 52, 53...
                    if (i % 3) != 0:
                        continue
                    else:
                        individual_visit_record[i] = {"in_window": 0, "out_window": 0, "status": ''}
                        # write down previous window's month
                        prev_q = i-3
                        window_begin = (i) * 28 - 14
                        window_end = (i) * 31 + 14
                        individual_visit_record[i]["status"] = 'norm'
                        if covid_start_days <= window_begin and window_end <= covid_end_days:
                            individual_visit_record[i]["status"] = 'covid'
                        # Write down out-window visit to previous window
                        while visit_class[ID]['record'] and visit_class[ID]['record'][0] < window_begin:
                            if (i - 3) in individual_visit_record:
                                individual_visit_record[prev_q]["out_window"] += + 1
                            visit_class[ID]['record'].popleft()
                        while visit_class[ID]['record'] and window_begin <= visit_class[ID]['record'][0] <= window_end:
                            if individual_visit_record[i]["in_window"] == 0:
                                individual_visit_record[i]["in_window"] += 1
                            visit_class[ID]['record'].popleft()

            visit_count[ID] = individual_visit_record


        # Same treatment for New ID's
        # The only difference is window's boundary calculation method
        else:
            for i in (range(total_weeks)):
                if i <=6:
                    individual_visit_record[i] = {"in_window": 0, "out_window": 0, "status": ''}
                    window_begin = (i)*30-7
                    window_end = (i)*30+7
                    individual_visit_record[i]["status"] = 'norm'
                    if covid_start_days <= window_begin and window_end <= covid_end_days:
                        individual_visit_record[i]["status"] = 'covid'

                    while visit_class[ID]['record'] and visit_class[ID]['record'][0] < window_begin:
                        if i > 0:
                            individual_visit_record[i-1]["out_window"] +=1
                        visit_class[ID]['record'].popleft()
                    while visit_class[ID]['record'] and window_begin <= visit_class[ID]['record'][0] <=window_end:
                        if individual_visit_record[i]["in_window"] == 0:
                            individual_visit_record[i]["in_window"] += 1
                        visit_class[ID]['record'].popleft()
                    if i == 0:
                        individual_visit_record[i]["in_window"] = 1
                        continue
                # When month > 6, new cohort ID requires to visit every 3 month
                else:
                    # If the month > 6, before the required visit month(9, 12...), skip these month. eg. 7, 8, 10, 11...
                    if (i%3) != 0:
                        continue
                    else:
                        individual_visit_record[i] = {"in_window": 0, "out_window": 0, "status": ''}
                        prev_q = i - 3
                        window_begin = i*30-14
                        window_end = i*30+14
                        individual_visit_record[i]["status"] = 'norm'
                        if covid_start_days <= window_begin and window_end <= covid_end_days:
                            individual_visit_record[i]["status"] = 'covid'
                        while visit_class[ID]['record'] and visit_class[ID]['record'][0] < window_begin:
                            if (i - 3) in individual_visit_record:
                                individual_visit_record[prev_q]["out_window"] +=1
                            visit_class[ID]['record'].popleft()
                        while visit_class[ID]['record'] and window_begin <= visit_class[ID]['record'][0] <=window_end:
                            if individual_visit_record[i]["in_window"] == 0:
                                individual_visit_record[i]["in_window"] += 1
                            visit_class[ID]['record'].popleft()
            # record each ID's all visits records in all their months in visit_count
            visit_count[ID] = individual_visit_record

    return visit_count

# Save visit_count to xlsx
def visit_count_to_excel(date, cohort_type, output_address):
    visit_count = calculation(date,cohort_type)
    records = []
    for pid, months in visit_count.items():
        for month, info in months.items():
            records.append({
                "ID": pid,
                "month": month,
                "in_window": info.get("in_window", 0),
                "out_window": info.get("out_window", 0),
                "status": info.get("status", "")
            })

    df = pd.DataFrame(records)
    df.to_excel(output_address, index=False)
    print("✅ Excel saved successfully.")

# ----------------------------------------------------------------------
# 7. Monthly aggregation and summary
# ----------------------------------------------------------------------
def count_output(current_time,cohort_type):
    """
    Aggregate compliance statistics for each study month.

    Returns
    -------
    month_count : dict
        {month: {
            'Visits Expected': int,
            'Visits Completed': int,
            'Completed %': float,
            'Visits Completed In Window': int,
            'In Window %': float
        }}
    """
    # Get 'Origin' or 'New' ID's visit_count
    visit_count = calculation(current_time,cohort_type)

    # Keep 'Origin' or 'New' monthly performance in dictionary
    #     month_count structure : dict
    #         {month: {
    #             'Visits Expected': int,
    #             'Visits Completed': int,
    #             'Completed %': float,
    #             'Visits Completed In Window': int,
    #             'In Window %': float
    #         }}
    month_count = {}

    # GO thorough each ID
    for key in visit_count.keys():
        # Go through each ID's each month record
        for month in visit_count[key].keys():
            # If current month record not in month_count, then initiate month's record
            if month not in month_count.keys():
                month_count[month] = {'Visits Expected':0,'Visits Completed':0,'Completed %':0,'Visits Completed In Window':0,'In Window %':0}
            # If current ID's current not belongs to 'covid' period
            if visit_count[key][month]['status'] == 'norm':
                # 'Visits Expected' add 1
                month_count[month]['Visits Expected']+=1
                # If there is 'in_window' record, both 'Visits Completed' and 'Visits Completed In Window' add 1
                if visit_count[key][month]['in_window'] > 0:
                    month_count[month]['Visits Completed'] += 1
                    month_count[month]['Visits Completed In Window'] += 1
                # If there is no 'in_window' record
                else:
                    # If there is 'out_window' record, only 'Visits Completed' add 1
                    if visit_count[key][month]['out_window'] > 0:
                        month_count[month]['Visits Completed'] += 1
            # For covid period situation
            else:
                # If there is visit record, no matter what visit type is.
                # Both 'Visits Expected' and 'Visits Completed' add 1, and this visit accounts for 'Visits Completed In Window'
                # If there is no visit record, everything keep what it is
                if visit_count[key][month]['in_window'] > 0 or visit_count[key][month]['out_window'] > 0:
                    month_count[month]['Visits Expected'] += 1
                    month_count[month]['Visits Completed'] += 1
                    month_count[month]['Visits Completed In Window'] += 1
    # After count each ID's monthly performance,
    # go through monly performance record again and calculate 'Completed %' and 'In Window %'
    for month in month_count.keys():
        if month_count[month]['Visits Expected'] != 0:
            month_count[month]['Completed %'] = float(month_count[month]['Visits Completed'] * 100 / month_count[month]['Visits Expected'])
            month_count[month]['In Window %'] = float(month_count[month]['Visits Completed In Window'] *100 / month_count[month]['Visits Expected'])

    return month_count, cohort_type


# ----------------------------------------------------------------------
# 8. Write results to Excel
# ----------------------------------------------------------------------
def to_excel(month_count, cohort_type):
    """
    Write summarized monthly compliance results to the correct Excel sheet.
    """
    # 1. Convert month_count dict → DataFrame
    df = pd.DataFrame.from_dict(month_count, orient='index')
    df.reset_index(inplace=True)
    df.rename(columns={'index': 'month'}, inplace=True)
    df = df.sort_values(by='month')

    # 2. Accodring to cohort type, write into sheet
    if cohort_type == 'origin':
        sheet_name = "2_Orig Cohort Sample Output"
    elif cohort_type == 'new':
        sheet_name = "3_New Cohort Sample Output"
    else:
        raise ValueError(f"❌ Unknown cohort_type: {cohort_type}")

    # 3. read original REACH Visit Compliance.xlsx and write output into corresponding sheet
    with pd.ExcelWriter(address, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"✅ Results for '{cohort_type}' cohort successfully written to '{sheet_name}'.")


if __name__ == "__main__":
    # Get both 'origin' and 'new' cohort output
    for c in ['origin', 'new']:
        month_count, cohort = count_output("10/13/2025", c)
        to_excel(month_count, cohort)



