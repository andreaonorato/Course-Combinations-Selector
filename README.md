# Course Combinations Selector

## What is it?

The **Course Combinations Selector** is a Python script designed to help students efficiently plan their course schedules. It generates all possible combinations of courses that sum up to a specified number of credits (e.g., 30 credits) while ensuring that the selected courses do not have overlapping schedules. The script also provides a summary of the total hours required for each combination, making it easier to balance your academic workload.

## How to use it?
1. **Prepare Your Course Data in Excel**:
   - Create an Excel file with the following columns:
     - `TYPE`: The course code (optional).
     - `NAME`: The name of the course.
     - `CREDITS`: The number of credits awarded for completing the course.
     - `EXAM TYPE`: The type of exam associated with the course (optional).
     - `LINK`: A link to the course page or resources (optional).
     - `TIME SCHEDULE`: The time schedule of the course, formatted as `Day Start-End` (e.g., `Mon 10-12`, `Thu 14-17`). If a course has multiple sessions, list them all in the same cell separated by a space.
     - `Difficulty(1-5)`: A self-assessed difficulty rating for the course, where 1 is easiest and 5 is most difficult (optional).
   
   - Fill in the Excel file with all your courses and their respective details.
   
   - Save the Excel file in `.xlsx` format. Ensure that the file is saved in the same directory as the script or update the script with the correct file path.
1. **Launch main.py**:
   - This will produce an Excel output file with all the possible combinations

### Prerequisites

- Python 3.x
- Required Python libraries: `pandas`, `openpyxl`

You can install the required libraries using pip:

```bash
pip install pandas openpyxl
