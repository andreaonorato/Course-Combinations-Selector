# Course Combinations Selector

## What is it?

The **Course Combinations Selector** is a Python script designed to help students efficiently plan their course schedules. It generates all possible combinations of courses that sum up to a specified number of credits (e.g., 30 credits) while ensuring that the selected courses do not have overlapping schedules. The script also provides a summary of the total hours required for each combination, making it easier to balance your academic workload.

## Features

- **Automatic Schedule Parsing**: Parses course schedules from an Excel file, handling multiple time slots per course.
- **Conflict-Free Combinations**: Ensures that course schedules do not overlap in the generated combinations.
- **Credit Summation**: Focuses on combinations that exactly meet the target credit requirement.
- **Total Hours Calculation**: Calculates the total hours required for each combination of courses.
- **User Comments Section**: Includes a section in the output Excel file for users to add their analysis and comments on each combination.

## How to Use

### Prerequisites

- Python 3.x
- Required Python libraries: `pandas`, `openpyxl`

You can install the required libraries using pip:

```bash
pip install pandas openpyxl
