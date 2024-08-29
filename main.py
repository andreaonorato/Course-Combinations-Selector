import pandas as pd
from itertools import combinations
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment

def time_to_minutes(time_str):
    """Convert time format 'x' or 'y' (hour) to minutes."""
    return int(time_str) * 60

def parse_schedule(schedule):
    """Parse a schedule string like 'Mon 10-12 Thu 14-17' into a dictionary."""
    days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    parsed_schedule = {day: [] for day in days}
    
    if pd.isna(schedule):
        return parsed_schedule
    
    sessions = schedule.split()
    for i in range(0, len(sessions), 2):
        day = sessions[i]
        start_time, end_time = sessions[i+1].split('-')
        start_time = time_to_minutes(start_time)
        end_time = time_to_minutes(end_time)
        parsed_schedule[day].append((start_time, end_time))
    
    return parsed_schedule

def calculate_total_hours(parsed_schedule):
    """Calculate the total number of hours for a given course based on its parsed schedule."""
    total_minutes = 0
    for day in parsed_schedule:
        for start_time, end_time in parsed_schedule[day]:
            total_minutes += (end_time - start_time)
    return total_minutes / 60  # Convert to hours

def schedules_overlap(schedule1, schedule2):
    """Check if two schedules overlap."""
    for day in schedule1:
        for time1 in schedule1[day]:
            for time2 in schedule2[day]:
                if max(time1[0], time2[0]) < min(time1[1], time2[1]):
                    return True
    return False

def find_combinations(courses, target_credits):
    """Find all course combinations that sum to target_credits and don't overlap in time."""
    valid_combinations = []
    for r in range(1, len(courses) + 1):
        for comb in combinations(courses, r):
            total_credits = sum(course['CREDITS'] for course in comb)
            if total_credits == target_credits:
                overlap = False
                for i in range(len(comb)):
                    for j in range(i + 1, len(comb)):
                        if schedules_overlap(comb[i]['parsed_schedule'], comb[j]['parsed_schedule']):
                            overlap = True
                            break
                    if overlap:
                        break
                if not overlap:
                    total_hours = sum(calculate_total_hours(course['parsed_schedule']) for course in comb)
                    valid_combinations.append((comb, total_hours))
    return valid_combinations

# Load data from Excel
file_path = 'courses epfl.xlsx'  # Update this with your file path
df = pd.read_excel(file_path)

# Parse the schedule and convert the dataframe into a list of dictionaries
courses = []
for _, row in df.iterrows():
    course = row.to_dict()
    course['parsed_schedule'] = parse_schedule(row['TIME SCHEDULE'])
    courses.append(course)

# Find all combinations with 30 credits that do not overlap
target_credits = 30
valid_combinations = find_combinations(courses, target_credits)

# Create a new Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Course Combinations"

# Write headers
headers = ["Combination Number", "Course Name", "Credits", "Time Schedule", "Total Hours", "Comments"]
ws.append(headers)

# Make headers bold and centered
for cell in ws[1]:
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center")

# Write each combination to the Excel file
row_number = 2
for i, (comb, total_hours) in enumerate(valid_combinations, 1):
    for course in comb:
        ws.append([i, course['NAME'], course['CREDITS'], course['TIME SCHEDULE'], total_hours, ""])
    row_number += len(comb)
    ws.append([])  # Add a blank row between combinations

# Auto-size columns based on content
for column in ws.columns:
    max_length = max(len(str(cell.value)) for cell in column)
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column[0].column_letter].width = adjusted_width

# Save the Excel workbook
output_file_path = "Course_Combinations.xlsx"
wb.save(output_file_path)

print(f"Excel file '{output_file_path}' created successfully.")
