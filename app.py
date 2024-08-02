import streamlit as st
import pandas as pd
import random
import copy
import io
from openpyxl import Workbook

# Original functions
def has_collision(timetable, day, period_range):
    return any(period < len(timetable[day]) and timetable[day][period] is not None for period in period_range)

def exceeds_daily_limit(timetable, day, subject):
    matching_elements = [x for x in labs if x in timetable[day]]
    if matching_elements:
        return timetable[day].count(subject) >= 1
    else:
        return timetable[day].count(subject) >= 2

pcs_assigned = []

def assign_continuous(timetable):
    for special, count in continues_count.items():
        assigned = False
        for day in random.sample(days, len(days)):
            available_periods = list(range(7 - count + 1))
            available_periods.remove(3)
            random.shuffle(available_periods)
            for period in available_periods:
                if not has_collision(timetable, day, range(period, period + count)) and (day, period) not in pcs_assigned:
                    for j in range(count):
                        timetable[day][period + j] = special
                        pcs_assigned.append((day, period))
                    assigned = True
                    break
            if assigned:
                break
        if not assigned:
            print(f"Warning: Unable to assign {special}")

def assign_others(timetable):
    library_assigned = False
    sports_assigned = False

    for day in random.sample(days, len(days)):
        if library_assigned and sports_assigned:
            break

        if not library_assigned:
            for period in [3, 6]:
                if not has_collision(timetable, day, [period]):
                    timetable[day][period] = 'Library'
                    library_assigned = True
                    break

        if not sports_assigned:
            if not has_collision(timetable, day, [6]):
                timetable[day][6] = 'Sports'
                sports_assigned = True

def assign_subjects(timetable_1, timetable_2):
    subject_counts = {subject: 0 for subject in subjects}
    temp_timetable = copy.deepcopy(timetable_1)
    temp_subs_count = subject_counts.copy()
    temp_subs = list(subject_counts.keys())
    for day in days:
        for period in range(len(periods) - 1):
            if temp_timetable[day][period] is not None:
                continue
            else:
                c = 0
                while (c <= 5):
                    c += 1
                    random.shuffle(temp_subs)
                    sub = random.choice(temp_subs)
                    if temp_subs_count[sub] >= subjects_frequency[sub]:
                        temp_subs.remove(sub)
                        continue
                    if timetable_2[day][period] is not None:
                        if faculties_section2[timetable_2[day][period]] == faculty_section1[sub]:
                            continue
                    if temp_timetable[day][period] is None and not exceeds_daily_limit(temp_timetable, day, sub):
                        temp_timetable[day][period] = sub
                        temp_subs_count[sub] += 1
                        break
                if c == 6:
                    return assign_subjects(timetable_1, timetable_2)
    return temp_timetable

def assign_subjects_section_2(timetable_1, timetable_2):
    subject_counts = {subject: 0 for subject in subjects}
    temp_timetable_2 = copy.deepcopy(timetable_2)
    temp_subs_count = subject_counts.copy()
    temp_subs = list(subject_counts.keys())
    for day in days:
        for period in range(len(periods) - 1):
            if temp_timetable_2[day][period] is not None:
                continue
            else:
                c = 0
                while (c <= 5):
                    c += 1
                    random.shuffle(temp_subs)
                    sub = random.choice(temp_subs)
                    if temp_subs_count[sub] >= subjects_frequency[sub]:
                        temp_subs.remove(sub)
                        continue
                    if timetable_1[day][period] is not None:
                        if faculty_section1[timetable_1[day][period]] == faculties_section2[sub]:
                            continue
                    if temp_timetable_2[day][period] is None and not exceeds_daily_limit(temp_timetable_2, day, sub):
                        temp_timetable_2[day][period] = sub
                        temp_subs_count[sub] += 1
                        break
                if c == 6:
                    return assign_subjects_section_2(timetable_1, timetable_2)
    return temp_timetable_2

def create_empty_timetable():
    return {day: [None] * 7 for day in days}

def insert_lunch_break(timetable):
    for day in days:
        timetable[day].insert(4, "Lunch Break")
        
def create_faculty_subject_timetable(timetable, faculty_dict):
    faculty_subject_timetable = copy.deepcopy(timetable)
    for day in days:
        for period in range(len(periods)):
            subject = timetable[day][period]
            if subject in faculty_dict:
                faculty_subject_timetable[day][period] = f"{subject} ({faculty_dict[subject]})"
    return faculty_subject_timetable

def create_faculty_timetables(timetable_section_1, timetable_section_2):
    faculty_timetables = {}
    
    all_faculties = {**faculty_section1, **faculties_section2}
    
    for subject, faculty in all_faculties.items():
        if faculty not in faculty_timetables:
            faculty_timetables[faculty] = create_empty_timetable()
        
        for day in days:
            period_index = 0
            for period in range(len(periods)):
                if periods[period] == 'Lunch Break':
                    continue
                if timetable_section_1[day][period] == subject:
                    faculty_timetables[faculty][day][period_index] = f"{subject} (Section 1)"
                elif timetable_section_2[day][period] == subject:
                    faculty_timetables[faculty][day][period_index] = f"{subject} (Section 2)"
                period_index += 1
    
    return faculty_timetables

# Streamlit app
st.title("Timetable Generator")

# Sidebar file upload
data_file = st.sidebar.file_uploader("Upload data_1.xlsx", type="xlsx")

if data_file:
    # Read data from Excel file with two sheets
    df_subjects = pd.read_excel(data_file, sheet_name="Sheet1")
    df_labs = pd.read_excel(data_file, sheet_name="Sheet2")

    # Extract data from DataFrame
    subjects = df_subjects['subjects'].dropna().tolist()
    labs = df_subjects['labs'].dropna().tolist()
    subjects_frequency = dict(zip(df_subjects['subjects'].dropna(), df_subjects['subjects_frequency'].dropna()))
    faculty_section1 = dict([item.split(':') for item in df_subjects['faculty_section1'].dropna().tolist()])
    faculties_section2 = dict([item.split(':') for item in df_subjects['faculty_section2'].dropna().tolist()])

    count = [2, 2]
    continuous = ['PCS - 1', 'PCS - 2']
    continues_count = {i: int(j) for i, j in zip(continuous, count)}

    # Initialize the timetable
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    periods = ['Period 1', 'Period 2', 'Period 3', 'Period 4', 'Lunch Break', 'Period 5', 'Period 6', 'Period 7']

    # Generate timetables
    timetable_section_1 = create_empty_timetable()
    timetable_section_2 = create_empty_timetable()

    # Assign lab periods from Excel file
    for index, row in df_labs.iterrows():
        day = row['Day']
        periods_range = list(range(int(row['Periods'].split(',')[0]), int(row['Periods'].split(',')[1])))
        lab = row['Lab']
        section = row['Section']
        if section == 1:
            for period in periods_range:
                timetable_section_1[day][period] = lab
        else:
            for period in periods_range:
                timetable_section_2[day][period] = lab

    # Assign continuous subjects for both sections
    assign_continuous(timetable_section_1)
    assign_continuous(timetable_section_2)

    # Assign Library and Sports for both sections
    assign_others(timetable_section_1)
    assign_others(timetable_section_2)

    # Assign subjects for Section 1
    timetable_section_1 = assign_subjects(timetable_section_1, timetable_section_2)

    # Assign subjects for Section 2
    timetable_section_2 = assign_subjects_section_2(timetable_section_1, timetable_section_2)

    # Insert lunch break
    insert_lunch_break(timetable_section_1)
    insert_lunch_break(timetable_section_2)

    # Create faculty-subject timetables for both sections
    faculty_subject_timetable_section_1 = create_faculty_subject_timetable(timetable_section_1, faculty_section1)
    faculty_subject_timetable_section_2 = create_faculty_subject_timetable(timetable_section_2, faculties_section2)
    # Create faculty timetables
    faculty_timetables = create_faculty_timetables(timetable_section_1, timetable_section_2)

    # Convert timetables to DataFrame for better display
    timetable_df_section_1 = pd.DataFrame(timetable_section_1, index=periods).T
    timetable_df_section_2 = pd.DataFrame(timetable_section_2, index=periods).T
    faculty_subject_df_section_1 = pd.DataFrame(faculty_subject_timetable_section_1, index=periods).T
    faculty_subject_df_section_2 = pd.DataFrame(faculty_subject_timetable_section_2, index=periods).T

    # Create Excel file
    def create_excel_file():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Write timetables to Excel
            timetable_df_section_1.to_excel(writer, sheet_name="Section 1 Timetable")
            timetable_df_section_2.to_excel(writer, sheet_name="Section 2 Timetable")
            faculty_subject_df_section_1.to_excel(writer, sheet_name="Section 1 Faculty-Subject")
            faculty_subject_df_section_2.to_excel(writer, sheet_name="Section 2 Faculty-Subject")

            for faculty, timetable in faculty_timetables.items():
                faculty_df = pd.DataFrame(timetable, index=[p for p in periods if p != 'Lunch Break'])
                faculty_df.T.to_excel(writer, sheet_name=f"{faculty} Timetable")

        processed_data = output.getvalue()
        return processed_data

    # Dropdown for timetable selection
    option = st.selectbox(
        "Select Timetable to View",
        ["Section 1 Timetable", "Section 2 Timetable", "Section 1 Faculty-Subject", "Section 2 Faculty-Subject"]
    )

    if option == "Section 1 Timetable":
        st.header("Timetable for Section 1")
        st.dataframe(timetable_df_section_1)

    elif option == "Section 2 Timetable":
        st.header("Timetable for Section 2")
        st.dataframe(timetable_df_section_2)

    elif option == "Section 1 Faculty-Subject":
        st.header("Faculty-Subject Timetable for Section 1")
        st.dataframe(faculty_subject_df_section_1)

    elif option == "Section 2 Faculty-Subject":
        st.header("Faculty-Subject Timetable for Section 2")
        st.dataframe(faculty_subject_df_section_2)

    # Display faculty timetables
    st.header("Individual Faculty Timetables")
    for faculty, timetable in faculty_timetables.items():
        st.write(f"Timetable for {faculty}:")
        faculty_df = pd.DataFrame(timetable, index=[p for p in periods if p != 'Lunch Break'])
        st.dataframe(faculty_df)
        st.write("=" * 50)

    # Download button
    excel_file = create_excel_file()
    st.download_button(
        label="Download Timetables as Excel",
        data=excel_file,
        file_name="timetables.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("Please upload data_1.xlsx file to generate timetables.")    
