# Myanmar XLSX to CSV

#### This Python script grabs data from the specified xlsx that the user inputs and generates a csv file called `myanmar_clean_data.csv` which can then be read by R to create RShiny Dashboards.

###Install:

1. Install Python 2.7.10: https://www.python.org/downloads/
2. Install Pip 
*More information found here for your specific distribution: http://docs.python-guide.org/en/latest/
3. Get the code: git clone https://github.com/rebeldroid12/myanmar_xlsx_to_csv.git
4. Install project requirements
5. Run the program!
	- open up the command line
	- navigate to the script `myanmar_xlsx_to_csv.py`
	- find the path to desired .xlsx file
	- Run Python script with desired .xlsx file
	- Look for `myanmar_clean_data.csv` in your current directory
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
python myanmar_xlsx_to_csv.py C:\Users\Desktop\myfile.xlsx
```
It will print out 'All done!' when the .csv file is created. Enjoy the clean data :)

