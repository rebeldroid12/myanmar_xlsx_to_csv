# Myanmar XLSX to CSV

#### This Python script grabs data from the specified xlsx that the user inputs and generates a csv file called `YYYY-YY_myanmar_clean_data.csv` (YYYY-YY is the year the user specified the data was from) which can then be read by R to create RShiny Dashboards.

###Install:

1. Install Python 2.7.10: https://www.python.org/downloads/
2. Install Pip 
*More information found here for your specific distribution: http://docs.python-guide.org/en/latest/
3. Get the code: git clone https://github.com/rebeldroid12/myanmar_xlsx_to_csv.git
4. Install project requirements
5. Run the program!
	- open up the command line
	- navigate to the script `myanmar_xlsx_to_csv_user_input.py`
	- find the path to desired .xlsx file
	- Run Python script with desired .xlsx file
	- Look for `YYYY-YY_myanmar_clean_data.csv` in your `\clean_data\` directory
	- Gaze at the clean data :)


###Example for Ubuntu:
(1): Installing Python

```
sudo apt-get install python2.7
sudo apt-get install python-dev
```

(2): Install pip
```
sudo apt-get install python-pip
```

(3): Get the code
```
git clone https://github.com/rebeldroid12/myanmar_xlsx_to_csv.git
```

(4): Install project requirements
```
pip install -r requirements.txt
```

(5): Running the program. Once you have opened the command line and navigated to the directory called `myanmar_xlsx_to_csv`, locate the xlsx you would like to run and grab the path for it. i.e. `C:\Users\Desktop\myfile.xlsx`. Go to your command line:

*You should be in the directory `myanmar_xlsx_to_csv` or know the path to the script to be able to call it*
```
python myanmar_xlsx_to_csv_user_input.py C:\Users\Desktop\myfile.xlsx
```
It will print out:

```
Data can be found in *YYYY-YY*_myanmar_clean_data.csv  
Created by Loren Velasquez as part of Statistics Without Borders. 
Code is open sourced and found at https://github.com/rebeldroid12/myanmar_xlsx_to_csv
```
---
---
---

# XLSX FILE:

## Make sure...

1. Column A does not have any information

2. Each sheet tab is the name of the corresponding Region the worksheet is about 

3. Cell B1 the first word is the Region Name

4. Flows are either Income or Expenditure and are found in column C

5. Entities are either (High Court, Advocate General, Auditor General) (In million kyats), (Ministries, Administrative Departments, Municipals) (in Million kyats), or (State Owned Enterprises) (in million kyats) and are found in column B

6. Budgets are between the rows labeled `Budget item` and `Total` and are found in column B

7. Sources are in the same line as 'Budget Item' (from column B) and are found in column C

8. Values of interest encompass the data between `Budget Items` and `Total` in the B column (but not including those rows), up to the Total column in column L (but not including those rows)


Please refer to [sample.xlsx](https://github.com/rebeldroid12/myanmar_xlsx_to_csv/blob/master/sample.xlsx) for the file format needed in order for this program to run correctly.
