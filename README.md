# Gantt Chart Monthly Schedule Generateor

## Introduction
This script creates a gantt chart from an excel file that contains tasks and their start and end dates. It is intended for monthly schedules that have recurring tasks every month at the same date.

## Requirements
  - Python 3.11.9
  - Pandas 1.5.3
  - openpyxl 3.1.2

## Details
This script:
1. Takes in 'monthly-schedule.xlsx' and user input (MM-YYYY) as inputs. 
2. Creates a df consisting of columns from 'monthly-schedule.xlsx' and adds a gantt chart next to it. The gantt chart corresponds to the tasks and their start and end date. 
3. Generates an excel file from the df. It also marks dates that are weekends (red color). 
4. The excel file formatting from the 'monthly-schedule.xlsx' is preserved in the output excel file
5. The output excel file is "Monthly Timeline <Month> <Year>.xlsx". The filename changes according to the month and year from the user input. 

## Notes
  - The script does not include tests for the user input (MM-YYYY). Make sure to input the month and year correctly according to the format
  - The 'monthly-schedule.xlsx' is customizable. HOWEVER, make sure the first column is the task and make sure that columns 'start_date' and 'end_date' exist. The 'Person in Charge' column is just an addition. It can be changed with other columns or completely removed.