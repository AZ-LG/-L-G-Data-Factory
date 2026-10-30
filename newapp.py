"""Weekly Updates app, adapted from Code_Revision_v3.ipynb.
Run locally: python -m streamlit run newapp.py
"""
import io
import re
from datetime import date, datetime
from copy import copy
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def read_table(workbook_bytes, sheet_name, required):
    """Detect a report header by labels, including header-only reports."""
    raw = pd.read_excel(io.BytesIO(workbook_bytes), sheet_name=sheet_name, header=None)
    text = raw.apply(lambda col: col.astype("string").str.strip())
    matches = pd.Series(True, index=raw.index)
    for label in required:
        matches &= text.eq(label).any(axis=1)
    if not matches.any():
        raise ValueError(f"Cannot find headers {required} in '{sheet_name}'.")
    header_row = matches[matches].index[0]
    headers = raw.iloc[header_row].map(lambda v: v.strip() if isinstance(v, str) else v)
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df.columns.name = None
    df = df.loc[:, df.columns.notna() & (df.columns != "")]
    if df.columns.duplicated().any():
        raise ValueError(f"Duplicate headers in '{sheet_name}'.")
    return df.reset_index(drop=True)


def get_site_master(students_bytes):
    hours = read_table(students_bytes, 'Participant Attendance Hours', ['Site', 'Institution'])
    sites = hours['Site'].astype('string').str.strip()
    valid = sites.notna() & sites.ne('') & ~sites.str.lower().isin(['nan', 'total', 'subtotal'])
    result = sites.loc[valid].drop_duplicates().rename('Site').to_frame().reset_index(drop=True)
    if result.empty:
        raise ValueError('No sites found in student Participant Attendance Hours.')
    return result


def clean_session_rows(df, site_master, report_name):
    """Normalize merge keys and check grant sites without editing uploads."""
    for col in ["Site", "Activity", "Session"]:
        df[col] = df[col].astype("string").str.strip()
    df = df.loc[
        df["Site"].notna() & df["Site"].ne("")
        & ~df["Site"].str.lower().isin(["total", "subtotal", "average"])
    ].copy()
    unexpected = set(df["Site"]) - set(site_master["Site"])
    if unexpected:
        raise ValueError(f"Sites missing from master list in {report_name}: {unexpected}")
    return df.reset_index(drop=True)


from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from copy import copy

def clean_sheet_name(name):
    name = str(name).strip()
    for character in [':', '/', '\\', '?', '*', '[', ']']:
        name = name.replace(character, '')
    return name.strip("'")[:31] or 'Sheet'

def apply_missing_highlights(output_filename, total_missing_rows, flag_cols, young_dob_rows, staff_missing_rows=None, staff_flag_cols=None, site_tables=None):
    red_fill = PatternFill('solid', start_color='FF9999', end_color='FF9999')
    blue_fill = PatternFill('solid', start_color='9999FF', end_color='9999FF')
    wb = load_workbook(output_filename)
    flag_to_original = {flag: flag.removesuffix('_missing') for flag in flag_cols if flag.endswith('_missing') and flag != 'State ParticipantID_missing'}
    for site, group in total_missing_rows.groupby('Site'):
        sheet_name = clean_sheet_name('Missing - ' + str(site))
        ws = wb[sheet_name]
        headers = {cell.value: cell.column for cell in ws[1]}
        for row_idx, (_, row) in enumerate(group.iterrows(), start=2):
            for flag_col, original_col in flag_to_original.items():
                if original_col in headers and row.get(flag_col, 0) == 1:
                    ws.cell(row=row_idx, column=headers[original_col]).fill = red_fill
    if staff_missing_rows is not None and staff_flag_cols is not None:
        ws = wb['Pull out - Missing Staff Info']
        headers = {cell.value: cell.column for cell in ws[1]}
        staff_mapping = {'Staff Type_missing': 'Staff Type', 'Employment Type_missing': 'Compensation Type', 'Funded by 21st CCLC_missing': 'Funder'}
        for row_idx, (_, row) in enumerate(staff_missing_rows.iterrows(), start=2):
            for flag_col in staff_flag_cols:
                original_col = staff_mapping[flag_col]
                if original_col in headers and row.get(flag_col, 0) == 1:
                    ws.cell(row=row_idx, column=headers[original_col]).fill = red_fill
    ws = wb['DOB - Atypical Age Check']
    headers = {cell.value: cell.column for cell in ws[1]}
    if 'Date Of Birth' in headers:
        for row_idx in range(2, ws.max_row + 1):
            ws.cell(row=row_idx, column=headers['Date Of Birth']).fill = blue_fill
    if site_tables is not None:
        for site_name in site_tables:
            ws = wb[clean_sheet_name(site_name)]
            headers = {cell.value: cell.column for cell in ws[1]}
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    column_name = ws.cell(row=1, column=cell.column).value
                    is_blank = cell.value is None or str(cell.value).strip() in ['', '-']
                    assignment_missing = column_name in ['Staff Assigned', 'Partner Assigned'] and str(cell.value).strip().upper() == 'N'
                    if (is_blank and column_name != "Dates Missing Attendance") or assignment_missing:
                        cell.fill = red_fill
                count_col = "Days Missing Attendance to Date"
                if count_col in headers:
                    count = ws.cell(row=row[0].row, column=headers[count_col]).value
                    try:
                        has_missing = float(count) > 0
                    except (TypeError, ValueError):
                        has_missing = False
                    if has_missing:
                        for name in [count_col, "Missing Attendance (%)", "Dates Missing Attendance"]:
                            if name in headers:
                                ws.cell(row=row[0].row, column=headers[name]).fill = red_fill
    output_filename.seek(0)
    output_filename.truncate(0)
    wb.save(output_filename)
    wb.close()

def build_report(students_bytes, adults_bytes, all_bytes, target_values, progress=None,
                 annual_schedule_bytes=None, missing_attendance_bytes=None):
    """Process uploaded workbook bytes; return the final XLSX bytes."""
    if progress is None:
        progress = lambda message: None

    if annual_schedule_bytes is None or missing_attendance_bytes is None:
        raise ValueError("Upload both Full-Year Activity/Session Details and Gary's Missing Attendance Report.")

    site_master = get_site_master(students_bytes)
    if len(site_master) != len(target_values):
        raise ValueError("Number of targets must match the site list.")
    if any(value <= 0 for value in target_values):
        raise ValueError("Targets must be positive.")
    site_master["Target # of students served"] = target_values

    progress("Processing Student Summary Statistics…")

    hour_sheet = next((name for name in pd.ExcelFile(io.BytesIO(students_bytes)).sheet_names if "Participants By Hour Band" in name), None)
    if hour_sheet is None:
        raise ValueError("Participants By Hour Band sheet not found.")
    df_part_by_hour = read_table(students_bytes, hour_sheet, ["Site", "Registered Youth"])

    site_names = df_part_by_hour['Site'].astype('string').str.strip()
    valid_site = site_names.notna() & site_names.ne('') & ~site_names.str.lower().isin(['nan', 'subtotal', 'total'])
    df_part_by_hour = df_part_by_hour.loc[valid_site].copy()
    df_part_by_hour['Site'] = site_names.loc[valid_site]
    df_part_by_hour = df_part_by_hour.reset_index(drop=True)

    unexpected_sites = set(df_part_by_hour["Site"]) - set(site_master["Site"])
    if unexpected_sites:
        raise ValueError(f"Hour-band sites missing from master list: {unexpected_sites}")

    daily_site_att = read_table(students_bytes, "Daily Site Attendance Summary", ["Site"])
    daily_site_att["Site"] = daily_site_att["Site"].astype("string").str.strip()
    daily_site_att = daily_site_att.loc[daily_site_att["Site"].isin(site_master["Site"])].copy()
    if daily_site_att["Site"].duplicated().any():
        raise ValueError("Duplicate sites found in daily attendance.")

    from datetime import datetime, date
    date_cols = [col for col in daily_site_att.columns if isinstance(col, (pd.Timestamp, datetime, date))]
    if daily_site_att.empty:
        daily_site_att['Total'] = pd.Series(dtype='Int64')
    else:
        if not date_cols:
            raise ValueError('No date columns found in daily attendance.')
        daily_values = daily_site_att[date_cols].apply(pd.to_numeric, errors='coerce')
        daily_site_att['Total'] = daily_values.mean(axis=1).round().astype('Int64')
    daily_site_att = daily_site_att[['Site', 'Total']]

    all_cols = ['0 Hours', 'Less Than 15 Hours', '15-44 Hours', '45-89 Hours', '90-179 Hours', '180-269 Hours', '270+ Hours']
    served_cols = ['Less Than 15 Hours', '15-44 Hours', '45-89 Hours', '90-179 Hours', '180-269 Hours', '270+ Hours']
    plus15_cols = ['15-44 Hours', '45-89 Hours', '90-179 Hours', '180-269 Hours', '270+ Hours']
    plus90_cols = ['90-179 Hours', '180-269 Hours', '270+ Hours']
    existing_cols = df_part_by_hour.columns.tolist()
    all_cols = [c for c in all_cols if c in existing_cols]
    served_cols = [c for c in served_cols if c in existing_cols]
    plus15_cols = [c for c in plus15_cols if c in existing_cols]
    plus90_cols = [c for c in plus90_cols if c in existing_cols]
    df_part_by_hour[all_cols] = df_part_by_hour[all_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
    df_part_by_hour['Registered Youth'] = pd.to_numeric(df_part_by_hour['Registered Youth'], errors='coerce').fillna(0)
    df_totals = pd.DataFrame({'Site': df_part_by_hour['Site'], 'Total # Enrolled': df_part_by_hour['Registered Youth'], 'Total # Served': df_part_by_hour[served_cols].sum(axis=1), 'Total # 15+': df_part_by_hour[plus15_cols].sum(axis=1), 'Total # 90+': df_part_by_hour[plus90_cols].sum(axis=1)})
    df_totals = site_master[['Site', 'Target # of students served']].merge(df_totals, on='Site', how='left', validate='one_to_one')

    daily_site_att['Total'] = daily_site_att['Total'].astype('Int64')
    daily_site_att = daily_site_att.rename(columns={'Total': 'Avg. # of Students Per Day'})
    df_totals = df_totals.merge(daily_site_att[['Site', 'Avg. # of Students Per Day']], on='Site', how='left', validate='one_to_one')

    df_totals = df_totals.rename(columns={'Site': 'School'})
    df_totals = df_totals[['School', 'Target # of students served', 'Total # Enrolled', 'Total # Served', 'Avg. # of Students Per Day', 'Total # 15+', 'Total # 90+']]
    total_row = df_totals.iloc[:, 1:].sum(min_count=1).to_frame().T
    total_row.insert(0, 'School', 'Total')
    if df_totals['Avg. # of Students Per Day'].isna().any():
        total_row['Avg. # of Students Per Day'] = pd.NA
    df_totals = pd.concat([df_totals, total_row], ignore_index=True)

    for hours in [15, 90]:
        count_col = f'Total # {hours}+'
        display_col = f'# of students {hours}+ hrs total (% of Target)'
        if count_col in df_totals.columns:
            counts = pd.to_numeric(df_totals[count_col], errors='coerce').astype('Int64')
            targets = pd.to_numeric(df_totals['Target # of students served'], errors='coerce').astype('Float64')
            percentages = counts.div(targets.where(targets > 0)).mul(100).round().astype('Int64')
            df_totals[display_col] = counts.astype('string').fillna('-') + ' (' + percentages.astype('string').fillna('-') + '%)'
            df_totals.loc[counts.isna(), display_col] = '-'
        elif display_col not in df_totals.columns:
            raise ValueError(f'Missing {count_col}. Rerun the student-count cells first.')
    df_totals = df_totals.drop(columns=['Total # 15+', 'Total # 90+'], errors='ignore').reset_index(drop=True)
    count_columns = ['Target # of students served', 'Total # Enrolled', 'Total # Served', 'Avg. # of Students Per Day']
    for col in count_columns:
        df_totals[col] = pd.to_numeric(df_totals[col], errors='coerce').astype('Int64')

    progress("Processing Family Component…")

    df_hours = pd.read_excel(io.BytesIO(adults_bytes), sheet_name='Participant Attendance Hours', header=None)
    header_text = df_hours.apply(lambda col: col.astype('string').str.strip())
    header_matches = header_text.eq('Site').any(axis=1) & header_text.eq('Participant Id').any(axis=1) & header_text.eq('Hours Present').any(axis=1)
    if not header_matches.any():
        raise ValueError('Cannot find Site, Participant Id, and Hours Present headers.')
    header_row = header_matches[header_matches].index[0]
    df_hours.columns = header_text.iloc[header_row]
    df_hours = df_hours.iloc[header_row + 1:].reset_index(drop=True)
    df_hours.columns.name = None

    df_hours['Site'] = df_hours['Site'].astype('string').str.strip()
    df_hours['Hours Present'] = pd.to_numeric(df_hours['Hours Present'], errors='coerce')
    df_hours['Participant Id'] = df_hours['Participant Id'].astype('string').str.strip().str.replace('\\.0$', '', regex=True)
    valid_id = df_hours['Participant Id'].notna() & df_hours['Participant Id'].ne('') & ~df_hours['Participant Id'].str.lower().isin(['nan', 'none'])
    df_active = df_hours.loc[(df_hours['Hours Present'] > 0) & valid_id & df_hours['Participant Id'].str.len().ne(9)].copy()
    if df_active['Site'].isna().any() or df_active['Site'].eq('').any():
        raise ValueError('An adult with attendance has no Site.')
    unexpected_sites = set(df_active['Site']) - set(site_master['Site'])
    if unexpected_sites:
        raise ValueError(f'Adult sites missing from master list: {unexpected_sites}')

    result = df_active.groupby('Site')['Participant Id'].nunique().reset_index(name='Parents Served (Total)')
    result = site_master[['Site']].merge(result, on='Site', how='left', validate='one_to_one')
    result['Parents Served (Total)'] = result['Parents Served (Total)'].fillna(0).astype(int)
    result.loc[len(result)] = {'Site': 'Total', 'Parents Served (Total)': result['Parents Served (Total)'].sum()}

    progress("Processing Participant Demographics…")

    df_part_demo = read_table(students_bytes, "Participant Demographics", ["Site", "Date Of Birth", "Grade Level"])
    required_columns = ['Site', 'Last Name', 'First Name', 'ParticipantID', 'State ParticipantID', 'Date Of Birth', 'Grade Level', 'Gender']
    missing_columns = [col for col in required_columns if col not in df_part_demo.columns]
    if missing_columns:
        raise ValueError(f'Missing demographics columns: {missing_columns}')
    identity_columns = ['Last Name', 'First Name', 'ParticipantID']
    has_identity = df_part_demo[identity_columns].apply(lambda col: col.astype('string').str.strip().fillna('').ne('')).any(axis=1)
    df_part_demo = df_part_demo.loc[has_identity].reset_index(drop=True)

    all_missing_flags = df_part_demo.copy()

    def clean_text(series):
        return series.astype('string').str.strip()

    def is_missing(series):
        text = clean_text(series)
        return text.isna() | text.str.lower().isin(['', 'not entered', 'nan', 'none'])
    all_missing_flags['Site'] = clean_text(all_missing_flags['Site'])
    all_missing_flags.loc[is_missing(all_missing_flags['Site']), 'Site'] = pd.NA
    for col in ['ParticipantID', 'State ParticipantID']:
        all_missing_flags[col] = clean_text(all_missing_flags[col]).str.replace('\\.0$', '', regex=True)
    dob_parsed = pd.to_datetime(all_missing_flags['Date Of Birth'], errors='coerce')
    all_missing_flags['Date Of Birth_missing'] = dob_parsed.isna().astype(int)
    all_missing_flags['ParticipantID_missing'] = (~all_missing_flags['ParticipantID'].str.fullmatch('[12]\\d{8}', na=False)).astype(int)
    all_missing_flags['Grade Level_missing'] = is_missing(all_missing_flags['Grade Level']).astype(int)
    all_missing_flags['Gender_missing'] = (~clean_text(all_missing_flags['Gender']).str.lower().isin(['male', 'female', 'non-binary'])).astype(int)
    all_missing_flags['State ParticipantID_missing'] = (~all_missing_flags['State ParticipantID'].str.fullmatch('\\d{10}', na=False)).astype(int)

    report_date = pd.Timestamp.now("America/New_York").normalize().tz_localize(None)
    age_tolerance = 3
    grade_text = clean_text(all_missing_flags['Grade Level']).str.upper()
    grade_number = pd.to_numeric(grade_text, errors='coerce')
    is_pk_kg = grade_text.isin(['PK', 'PRE-K', 'PREK', 'PRE-KINDERGARTEN', 'K', 'KG', 'KN', 'KINDERGARTEN']) | grade_number.eq(0)
    recognized_grade = (grade_number.between(1, 12) & grade_number.mod(1).eq(0)).fillna(False)
    typical_age = (grade_number + 5).where(recognized_grade)
    birthday_not_reached = (dob_parsed.dt.month > report_date.month) | (dob_parsed.dt.month == report_date.month) & (dob_parsed.dt.day > report_date.day)
    age = report_date.year - dob_parsed.dt.year - birthday_not_reached.astype(int)
    pk_kg_mismatch = (is_pk_kg & dob_parsed.notna() & (age.lt(1) | age.gt(9))).fillna(False)
    age_grade_mismatch = (recognized_grade & dob_parsed.notna() & (age - typical_age).abs().gt(age_tolerance)).fillna(False)
    future_dob = dob_parsed.gt(report_date)
    all_missing_flags['DOB_age_grade_review'] = (pk_kg_mismatch | age_grade_mismatch | future_dob).astype(int)
    dob_review_details = pd.DataFrame({'Age at Report Date': age.astype('Int64'), 'Typical Age for Grade': typical_age.astype('Int64'), 'Review Reason': np.select([future_dob.to_numpy(dtype=bool), pk_kg_mismatch.to_numpy(dtype=bool), age_grade_mismatch.to_numpy(dtype=bool)], ['DOB is in the future', 'PK/KG age outside 1–9 years', 'Age more than 3 years from typical grade age'], default='')})
    unrecognized_grades = grade_text.loc[~recognized_grade & ~is_pk_kg & ~is_missing(all_missing_flags['Grade Level'])].drop_duplicates().tolist()

    summary_flags = ['Date Of Birth_missing', 'ParticipantID_missing', 'Grade Level_missing', 'Gender_missing', 'State ParticipantID_missing']
    unexpected_sites = set(all_missing_flags['Site'].dropna()) - set(site_master['Site'])
    if unexpected_sites:
        raise ValueError(f'Demographics sites missing from master list: {unexpected_sites}')
    missing_summary = all_missing_flags.groupby('Site')[summary_flags].sum().reset_index()
    missing_summary = site_master[['Site']].merge(missing_summary, on='Site', how='left', validate='one_to_one')
    missing_summary[summary_flags] = missing_summary[summary_flags].fillna(0).astype(int)
    total_row = {'Site': 'Total'}
    total_row.update(missing_summary[summary_flags].sum().to_dict())
    missing_summary = pd.concat([missing_summary, pd.DataFrame([total_row])], ignore_index=True)
    missing_summary = missing_summary.rename(columns={'Date Of Birth_missing': 'DOB_missing', 'State ParticipantID_missing': '10digit_State ParticipantID_missing'})

    flag_cols = ['Date Of Birth_missing', 'ParticipantID_missing', 'Grade Level_missing', 'Gender_missing']
    total_missing_rows = all_missing_flags.loc[all_missing_flags[flag_cols].sum(axis=1) > 0].copy()
    missing_site_rows = all_missing_flags.loc[all_missing_flags['Site'].isna()].copy()
    review_mask = all_missing_flags['DOB_age_grade_review'].eq(1)
    young_dob_rows = all_missing_flags.loc[review_mask].copy()
    young_dob_rows = young_dob_rows.join(dob_review_details.loc[review_mask])

    progress("Processing Staff Details…")

    df_staff = pd.read_excel(io.BytesIO(all_bytes), sheet_name='Staff Details', header=None)
    header_text = df_staff.apply(lambda col: col.astype('string').str.strip())
    header_matches = header_text.eq('Site').any(axis=1) & header_text.eq('Staff Type').any(axis=1) & header_text.eq('Compensation Type').any(axis=1)
    if not header_matches.any():
        raise ValueError('Cannot find the Staff Details header row.')
    header_row = header_matches[header_matches].index[0]
    df_staff.columns = header_text.iloc[header_row]
    df_staff = df_staff.iloc[header_row + 1:].reset_index(drop=True)
    df_staff.columns.name = None
    required_columns = ['Site', 'Email Address', 'First Name', 'Last Name', 'Staff Type', 'Compensation Type', 'Funder']
    missing_columns = [col for col in required_columns if col not in df_staff.columns]
    if missing_columns:
        raise ValueError(f'Missing staff columns: {missing_columns}')

    df = df_staff.copy()
    text_columns = ['Site', 'Email Address', 'First Name', 'Last Name', 'Staff Type', 'Compensation Type', 'Funder']
    for col in text_columns:
        df[col] = df[col].astype('string').str.strip().fillna('')
        df.loc[df[col].str.lower().isin(['nan', 'none', 'not entered']), col] = ''
    has_identity = df[['Email Address', 'First Name', 'Last Name']].ne('').any(axis=1)
    df = df.loc[has_identity].copy()
    if df['Site'].eq('').any():
        raise ValueError('A staff record has no Site assigned.')
    unexpected_sites = set(df['Site']) - set(site_master['Site'])
    if unexpected_sites:
        raise ValueError(f'Staff sites missing from master list: {unexpected_sites}')
    email = df['Email Address'].str.lower()
    name_id = df['First Name'].str.lower() + '|' + df['Last Name'].str.lower()
    df['_staff_id'] = np.where(email.ne(''), 'email:' + email, 'name:' + name_id)
    df = df.drop_duplicates(subset=['Site', '_staff_id']).copy()
    staff_type = df['Staff Type'].str.lower()
    comp_type = df['Compensation Type'].str.lower()
    funder = df['Funder']
    df['Staff Type_missing'] = staff_type.isin(['', 'other']).astype(int)
    df['Employment Type_missing'] = (~comp_type.isin(['paid', 'volunteer'])).astype(int)
    is_volunteer = comp_type.eq('volunteer')
    funder_blank = funder.eq('')
    funder_not_21cclc = ~funder.str.contains('21', case=False, na=False)
    df['Funded by 21st CCLC_missing'] = (~is_volunteer & funder_blank | ~funder_blank & funder_not_21cclc).astype(int)
    staff_flag_cols = ['Staff Type_missing', 'Employment Type_missing', 'Funded by 21st CCLC_missing']
    staff_missing_rows = df.loc[df[staff_flag_cols].sum(axis=1) > 0].copy()

    staff_missing_summary = df.groupby('Site').agg(**{'# of Active Program Staff': ('_staff_id', 'nunique'), 'Staff Type': ('Staff Type_missing', 'sum'), 'Employment Type': ('Employment Type_missing', 'sum'), 'Funded by 21st CCLC': ('Funded by 21st CCLC_missing', 'sum')}).reset_index()
    staff_missing_summary = site_master[['Site']].merge(staff_missing_summary, on='Site', how='left', validate='one_to_one')
    summary_columns = ['# of Active Program Staff', 'Staff Type', 'Employment Type', 'Funded by 21st CCLC']
    staff_missing_summary[summary_columns] = staff_missing_summary[summary_columns].fillna(0).astype(int)
    total_row = {'Site': 'Total'}
    total_row.update(staff_missing_summary[summary_columns].sum().to_dict())
    staff_missing_summary = pd.concat([staff_missing_summary, pd.DataFrame([total_row])], ignore_index=True)

    progress("Building Site Session Tables…")

    def read_session_report(sheet_name):
        raw = pd.read_excel(io.BytesIO(all_bytes), sheet_name=sheet_name, header=None)
        header_text = raw.apply(lambda col: col.astype('string').str.strip())
        header_matches = header_text.eq('Site').any(axis=1) & header_text.eq('Activity').any(axis=1) & header_text.eq('Session').any(axis=1)
        if not header_matches.any():
            raise ValueError(f"Cannot find Site, Activity, and Session headers in '{sheet_name}'.")
        header_row = header_matches[header_matches].index[0]
        headers = raw.iloc[header_row].map(lambda value: value.strip() if isinstance(value, str) else value)
        df = raw.iloc[header_row + 1:].copy()
        df.columns = headers
        df.columns.name = None
        df = df.loc[:, df.columns.notna() & (df.columns != '')]
        if df.columns.duplicated().any():
            raise ValueError(f"Duplicate headers in '{sheet_name}'.")
        for col in ['Site', 'Activity', 'Session']:
            df[col] = df[col].astype('string').str.strip()
        df = df.loc[df['Site'].notna() & df['Site'].ne('') & ~df['Site'].str.lower().isin(['total', 'subtotal', 'average'])].copy()
        unexpected_sites = set(df['Site']) - set(site_master['Site'])
        if unexpected_sites:
            raise ValueError(f"Sites missing from master list in '{sheet_name}': {unexpected_sites}")
        return df.reset_index(drop=True)
    df_act = read_session_report('Activity-Session Details')
    df_enr = read_session_report('Session Enrollment by Session')
    df_att = read_session_report('Daily Activity Attendance Summa')

    cols_act = ['Site', 'Activity', 'Session', 'Days Scheduled', 'Session Start Date', 'Designated Staff', 'Partners']
    df_act = df_act[cols_act].copy()
    df_act = df_act.rename(columns={'Days Scheduled': 'Days Scheduled to Date', 'Designated Staff': 'Staff Assigned', 'Partners': 'Partner'})
    for col in ['Staff Assigned', 'Partner']:
        df_act[col] = df_act[col].astype('string').str.strip()
        df_act.loc[df_act[col].str.lower().isin(['', 'not entered', 'nan', 'none']), col] = pd.NA
    df_act['Partner Assigned'] = np.where(df_act['Partner'].notna(), 'Y', 'N')
    df_act['Staff Assigned'] = np.where(df_act['Staff Assigned'].notna(), 'Y', 'N')
    today = pd.Timestamp.now("America/New_York").normalize().tz_localize(None)
    df_act["Days Scheduled to Date"] = pd.to_numeric(df_act["Days Scheduled to Date"], errors="coerce").astype("Int64")
    df_act["Session Start Date"] = pd.to_datetime(df_act["Session Start Date"], errors="coerce")

    cols_enr = ['Site', 'Activity', 'Session', 'Enrolled Count']
    df_enr = df_enr[cols_enr].copy()
    df_enr['Enrolled Count'] = pd.to_numeric(df_enr['Enrolled Count'], errors='coerce')
    df_enr = df_enr.rename(columns={'Enrolled Count': 'Enrolled Participant'})

    cols_att = ['Site', 'Activity', 'Session', 'Total']
    df_att = df_att[cols_att].copy()
    attendance_text = df_att['Total'].astype('string').str.replace('(?i)^average\\s*:\\s*', '', regex=True).str.strip()
    df_att['Total'] = pd.to_numeric(attendance_text, errors='coerce').round().astype('Int64')
    df_att = df_att.rename(columns={'Total': 'Average Daily Attendance'})

    # Full-year scheduled days are separate from the to-date denominator.
    merge_keys = ["Site", "Activity", "Session"]
    progress("Reading Full-Year Activity/Session Details…")
    df_annual = read_table(
        annual_schedule_bytes, "Activity-Session Details",
        merge_keys + ["Days Scheduled"]
    )[merge_keys + ["Days Scheduled"]].copy()
    df_annual = clean_session_rows(df_annual, site_master, "Full-Year Activity/Session Details")
    df_annual["Days Scheduled"] = pd.to_numeric(
        df_annual["Days Scheduled"], errors="coerce"
    ).astype("Int64")
    df_annual = df_annual.drop_duplicates()
    if df_annual.duplicated(merge_keys).any():
        raise ValueError("Conflicting annual scheduled-day counts for the same Site / Activity / Session.")
    df_act = df_act.merge(df_annual, on=merge_keys, how="left", validate="many_to_one", indicator="_annual_match")
    unmatched = df_act["_annual_match"].eq("left_only")
    if unmatched.any():
        progress(f"Warning: {len(df_act.loc[unmatched, merge_keys].drop_duplicates())} sessions have no matching full-year schedule; yearly Days Scheduled will be blank.")
    df_act = df_act.drop(columns="_annual_match")

    # Count unique missing dates and concatenate dates chronologically.
    progress("Processing Gary's Missing Attendance Report…")
    sheet_name = next(
        (name for name in pd.ExcelFile(io.BytesIO(missing_attendance_bytes)).sheet_names
         if "Missing Attendance" in name), None
    )
    if sheet_name is None:
        raise ValueError("Missing Attendance report sheet not found.")
    df_missing_att = read_table(
        missing_attendance_bytes, sheet_name, merge_keys + ["Date"]
    )[merge_keys + ["Date"]].copy()
    df_missing_att = clean_session_rows(df_missing_att, site_master, "Gary's Missing Attendance Report")
    if df_missing_att[merge_keys].isna().any().any() or df_missing_att[merge_keys].eq("").any().any():
        raise ValueError("Missing site, activity, or session in Gary's Missing Attendance Report.")
    df_missing_att["Date"] = pd.to_datetime(df_missing_att["Date"], errors="coerce").dt.normalize()
    if df_missing_att["Date"].isna().any():
        raise ValueError("Invalid or missing dates in Gary's Missing Attendance Report.")
    df_missing_att = df_missing_att.drop_duplicates(merge_keys + ["Date"]).sort_values(merge_keys + ["Date"])
    df_missing_att["Date Text"] = df_missing_att["Date"].dt.strftime("%m/%d/%Y")
    missing_attendance_summary = df_missing_att.groupby(merge_keys).agg(**{
        "Days Missing Attendance": ("Date", "nunique"),
        "Dates Missing Attendance": ("Date Text", lambda dates: "; ".join(dates)),
    }).reset_index()

    sites = site_master["Site"].tolist()
    df_act = df_act.drop_duplicates().reset_index(drop=True)
    for report_name, report_df in [
        ("Activity-Session Details", df_act),
        ("Session Enrollment by Session", df_enr),
        ("Daily Activity Attendance Summary", df_att),
    ]:
        if report_df.duplicated(merge_keys).any():
            raise ValueError(f"Review duplicate sessions in {report_name} before merging.")

    known_sessions = pd.concat([df_act[merge_keys], df_enr[merge_keys], df_att[merge_keys]], ignore_index=True).drop_duplicates()
    unmatched_missing = missing_attendance_summary.merge(known_sessions, on=merge_keys, how="left", indicator=True)
    unmatched_keys = unmatched_missing.loc[unmatched_missing["_merge"].eq("left_only"), merge_keys]
    if not unmatched_keys.empty:
        examples = "; ".join(" / ".join(str(row[col]) for col in merge_keys) for _, row in unmatched_keys.head(5).iterrows())
        raise ValueError(f"Missing-attendance sessions were not found in the session reports: {examples}. Check institution, academic year, report period, and session names.")

    progress("Building Site Session Tables…")
    site_tables = {}
    output_columns = [
        "Site", "Activity", "Session", "Days Scheduled",
        "Enrolled Participant", "Average Daily Attendance",
        "Missing Attendance (%)", "Staff Assigned", "Partner Assigned",
        "Days Scheduled to Date", "Days Missing Attendance to Date",
        "Dates Missing Attendance", "Partner",
    ]
    for site in sites:
        site_act = df_act.loc[df_act["Site"] == site].copy()
        site_enr = df_enr.loc[df_enr["Site"] == site].copy()
        site_att = df_att.loc[df_att["Site"] == site].copy()
        site_missing = missing_attendance_summary.loc[missing_attendance_summary["Site"] == site].copy()
        merged_site = site_act.merge(site_enr, on=merge_keys, how="outer", validate="one_to_one")
        merged_site = merged_site.merge(site_att, on=merge_keys, how="outer", validate="one_to_one")
        merged_site = merged_site.merge(site_missing, on=merge_keys, how="left", validate="one_to_one")
        merged_site = merged_site.loc[
            merged_site["Session Start Date"].isna()
            | merged_site["Session Start Date"].lt(today)
        ].copy()
        merged_site["Days Missing Attendance"] = merged_site["Days Missing Attendance"].fillna(0).astype("Int64")
        merged_site["Dates Missing Attendance"] = merged_site["Dates Missing Attendance"].fillna("-")
        for col in ["Staff Assigned", "Partner Assigned"]:
            merged_site[col] = merged_site[col].fillna("N")
        for col in ["Days Scheduled", "Days Scheduled to Date"]:
            merged_site[col] = pd.to_numeric(merged_site[col], errors="coerce").astype("Int64")
        denominator = merged_site["Days Scheduled to Date"]
        missing_pct = merged_site["Days Missing Attendance"].div(denominator.where(denominator > 0)).mul(100)
        merged_site["Missing Attendance (%)"] = (missing_pct.round().astype("Int64").astype("string") + "%").fillna("-")
        inconsistent = merged_site["Days Missing Attendance"].gt(denominator).fillna(False)
        if inconsistent.any():
            progress(f"Warning: {int(inconsistent.sum())} session(s) at {site} have more missing dates than scheduled days to date. Check report periods; percentages have not been capped.")
        merged_site = merged_site.rename(columns={"Days Missing Attendance": "Days Missing Attendance to Date"})
        merged_site = merged_site[output_columns].sort_values(["Activity", "Session"]).reset_index(drop=True)
        site_tables[site] = merged_site.astype(object).where(merged_site.notna(), "-")

    progress("Writing Excel Report…")
    output_filename = io.BytesIO()

    internal_columns = {'DOB_too_young', 'DOB_age_grade_review', '_staff_id'}
    hide_cols = {'Race/Ethnicity', 'English Learner Status', 'Lunch Status', 'Special Education Status', 'IDEA Disability Type'}
    student_display_cols = [col for col in total_missing_rows.columns if not col.endswith('_missing') and col not in internal_columns and (col not in hide_cols)]
    dob_display_cols = [col for col in young_dob_rows.columns if not col.endswith('_missing') and col not in internal_columns]
    staff_display_cols = [col for col in staff_missing_rows.columns if not col.endswith('_missing') and col not in internal_columns]
    sheet_names = ['Student Summary Statistics', 'Family Component Summary', 'Missing Student Summary', 'DOB - Atypical Age Check', 'Missing Staff Summary', 'Pull out - Missing Staff Info']
    if not missing_site_rows.empty:
        sheet_names.append('Pull out - Missing Site Info')
    sheet_names += [clean_sheet_name('Missing - ' + str(site)) for site in total_missing_rows['Site'].dropna().unique()]
    sheet_names += [clean_sheet_name(site) for site in site_tables]
    if len(sheet_names) != len(set((name.lower() for name in sheet_names))):
        raise ValueError('Some output sheet names become identical after cleaning or truncation. Resolve the names before exporting.')
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df_totals.to_excel(writer, sheet_name='Student Summary Statistics', index=False)
        result.to_excel(writer, sheet_name='Family Component Summary', index=False)
        missing_summary.to_excel(writer, sheet_name='Missing Student Summary', index=False)
        if not missing_site_rows.empty:
            missing_site_display_cols = [col for col in missing_site_rows.columns if not col.endswith('_missing') and col not in internal_columns and (col not in hide_cols)]
            missing_site_rows[missing_site_display_cols].to_excel(writer, sheet_name='Pull out - Missing Site Info', index=False)
        for site, group in total_missing_rows.groupby('Site'):
            group[student_display_cols].to_excel(writer, sheet_name=clean_sheet_name('Missing - ' + str(site)), index=False)
        young_dob_rows[dob_display_cols].to_excel(writer, sheet_name='DOB - Atypical Age Check', index=False)
        staff_missing_summary.to_excel(writer, sheet_name='Missing Staff Summary', index=False)
        staff_missing_rows[staff_display_cols].to_excel(writer, sheet_name='Pull out - Missing Staff Info', index=False)
        for site_name, final_df in site_tables.items():
            final_df.to_excel(writer, sheet_name=clean_sheet_name(site_name), index=False)
    apply_missing_highlights(output_filename, total_missing_rows, flag_cols, young_dob_rows, staff_missing_rows=staff_missing_rows, staff_flag_cols=staff_flag_cols, site_tables=site_tables)

    # Percentage font colors, without requiring notebook Styler/Jinja2.
    output_filename.seek(0)
    wb = load_workbook(output_filename)
    ws = wb['Student Summary Statistics']
    headers = {cell.value: cell.column for cell in ws[1]}
    col = headers['# of students 15+ hrs total (% of Target)']
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row, col)
        try:
            pct = int(str(cell.value).split('(')[1].replace('%)', '').strip())
        except (ValueError, IndexError):
            continue
        font = copy(cell.font)
        font.color = '008000' if pct >= 100 else 'FF0000'
        cell.font = font
    final_buffer = io.BytesIO()
    wb.save(final_buffer)
    wb.close()
    return final_buffer.getvalue()



def main():
    import streamlit as st

    st.set_page_config(
        page_title="Weekly Update Report Tool", page_icon="📊", layout="centered"
    )
    st.title("📊 Weekly Update Report Tool")
    st.markdown("Upload five Excel workbooks, confirm each site's target, and download your weekly report.")
    st.subheader("Step 1 · Upload Your Excel Files")
    columns = st.columns(3)
    with columns[0]:
        students_file = st.file_uploader("🎒 Students File", type=["xlsx"], key="students")
    with columns[1]:
        adults_file = st.file_uploader("👪 Adults File", type=["xlsx"], key="adults")
    with columns[2]:
        all_file = st.file_uploader("📋 Grant Level File", type=["xlsx"], key="all")

    extra_columns = st.columns(2)
    with extra_columns[0]:
        annual_file = st.file_uploader(
            "📅 Full-Year Activity/Session Details", type=["xlsx"], key="annual_schedule"
        )
    with extra_columns[1]:
        missing_file = st.file_uploader(
            "📝 Gary's Missing Attendance Report", type=["xlsx"], key="missing_attendance"
        )
    st.caption("Use the same institution and academic year for all five files. The Grant Level and Gary's reports must cover the same to-date reporting period; the full-year schedule must cover the entire academic year.")

    st.divider()
    institution_name = st.text_input("🏫 Institution Name", value="Institution")
    st.subheader("Step 2 · Set Target Enrollment per Site")
    st.caption("Sites come from the Students file's Participant Attendance Hours sheet, including site placeholders.")

    students_bytes = students_file.getvalue() if students_file is not None else None
    adults_bytes = adults_file.getvalue() if adults_file is not None else None
    all_bytes = all_file.getvalue() if all_file is not None else None
    annual_schedule_bytes = annual_file.getvalue() if annual_file is not None else None
    missing_attendance_bytes = missing_file.getvalue() if missing_file is not None else None
    target_values = []
    sites_ready = False
    if students_bytes is not None:
        try:
            sites = get_site_master(students_bytes)["Site"].tolist()
            for site in sites:
                value = st.number_input(
                    f"{site} — Target # of students", min_value=1, value=100,
                    step=1, key=f"target:{site}"
                )
                target_values.append(int(value))
            sites_ready = True
        except Exception as exc:
            st.error(f"Could not read the Students file: {exc}")
    else:
        st.info("Upload the Students file to show the sites and enter their targets.")

    ready = (sites_ready and adults_bytes is not None and all_bytes is not None
             and annual_schedule_bytes is not None and missing_attendance_bytes is not None)
    st.divider()
    st.subheader("Step 3 · Generate Report")
    if not ready:
        st.warning("Upload all five .xlsx files and check the site targets before generating.")

    # Invalidate an earlier report when uploads, targets, or institution change.
    import hashlib
    fingerprint = hashlib.sha256()
    for value in (students_bytes, adults_bytes, all_bytes, annual_schedule_bytes, missing_attendance_bytes):
        fingerprint.update(hashlib.sha256(value or b"").digest())
    fingerprint.update(repr((institution_name, target_values)).encode())
    current_key = fingerprint.hexdigest()
    if st.session_state.get("report_input_key") != current_key:
        st.session_state.pop("output_bytes", None)
        st.session_state.pop("report_input_key", None)
        st.session_state.pop("output_filename", None)

    if st.button("🚀 Generate Report", disabled=not ready, type="primary"):
        st.session_state.pop("output_bytes", None)
        stage = ["Reading uploaded workbooks"]
        with st.status("Running pipeline…", expanded=True) as status:
            def progress(message):
                stage[0] = message
                st.write(message)
            try:
                data = build_report(
                    students_bytes, adults_bytes, all_bytes,
                    target_values, progress=progress,
                    annual_schedule_bytes=annual_schedule_bytes,
                    missing_attendance_bytes=missing_attendance_bytes
                )
                st.session_state.output_bytes = data
                safe_institution = re.sub(r'[\\/:*?"<>|]', '_', institution_name.strip()).replace(' ', '_') or 'Institution'
                date_str = pd.Timestamp.now("America/New_York").strftime("%Y%m%d")
                st.session_state.output_filename = f"WeeklyUpdates_{safe_institution}_{date_str}.xlsx"
                st.session_state.report_input_key = current_key
                status.update(label="✅ Report ready!", state="complete")
            except Exception as exc:
                status.update(label="❌ Report could not be generated", state="error")
                st.error(f"{stage[0]}: {exc}")
                st.caption("Check the named input sheet and its headers. No report was generated from these inputs.")

    if st.session_state.get("output_bytes") is not None:
        st.success(f"Your report for {institution_name.strip() or 'Institution'} is ready.")
        st.download_button(
            "⬇️ Download report", data=st.session_state.output_bytes,
            file_name=st.session_state.output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )


if __name__ == "__main__":
    main()
