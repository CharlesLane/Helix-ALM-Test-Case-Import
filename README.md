# Helix-ALM-Test-Case-Import

## Purpose

The ConvertTC.py file will read an Excel file which contains test cases, and convert them into a .csv file in the markup language used by Helix ALM for import. Read more [here.](https://help.perforce.com/alm/helixalm/2020.1.0/client/Content/User/TCM/TestCaseTestRunMarkupCodes.htm?Highlight=test%20text)

## Formatting

There is one variable required to get started. It is on line 5 of ConvertTC.py, and it determines the path to your source Excel file.

Within your Excel file, your tests should be in the following format:

### Column 1: Test Case Name
This is a summary of the test case. Its name must be repeated exactly on each line.

### Column 2: Steps
This column contains the instructions for the tester to execute. Some lines may be blank, if you have multiple expected results as described in column 3.

### Column 3: Expected Results
This column contains the results you expect to see as a result of performing the action described in the associated step, located in column 2. Some lines may be blank, if the step in column 2 does not have an associated expected result.

**PLEASE NOTE**
The script will end when both the Steps and Expected results columns, columns 2 and 3, are blank.

### Column 4: Comments
Comments are optional additional information pertaining to the associated step and/or expected result.

### Columns 5+: Additional Fields
You may add additional columns which contain data attributes you would like to import into Helix ALM on the Test Case. The data on those additional fields can optionally be repeated on each line of the test case, or only on the first line of each new test case.

___

## Usage

1. Set the path to your Excel file on line 5 of ConvertTC.py
2. Save ConvertTC.py
3. Run the script
4. You'll find the resulting ConvertedTC.csv file in the same directory as ConvertTC.py
