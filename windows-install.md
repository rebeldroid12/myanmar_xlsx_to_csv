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


###Example for Windows:
(1): Installing Python

Go to https://www.python.org/downloads/ and download python2.7.11 or the latest version. 
- You will get a .msi file which installs Python for you. Python2.7 folder will get added to your computer; usually gets added in C:\Python27

Make sure to add Python to your path.
- Look up your environment variables 
![windows env vars](https://github.com/rebeldroid12/myanmar_xlsx_to_csv/blob/master/windows-resources/environment_vars.PNG)

- Edit the the User Path variable (at the top)
- Add the path ';C:\Python27'

![windows path var](https://github.com/rebeldroid12/myanmar_xlsx_to_csv/blob/master/windows-resources/user_path.PNG)

*semicolons separate the different packages/programs.*

(2): Install pip

All Python installations of Python 2.7.9+ come with pip so no need to install again.

(3): Get the code

Download the zip: https://github.com/rebeldroid12/myanmar_xlsx_to_csv

![download zip](https://github.com/rebeldroid12/myanmar_xlsx_to_csv/blob/master/windows-resources/download_zip.PNG)

Extract the zip file and remember where the folder will live (your directory path)

*One way to find out what directory you files live in is to use the Windows Explorer: *

![windows path dir](https://github.com/rebeldroid12/myanmar_xlsx_to_csv/blob/master/windows-resources/path_dir.PNG)

Here my directory path is: C:\Users\rebeldroid12\Downloads\myanmar_xlsx_to_csv-master\myanmar_xlsx_to_csv-master


(4): Install project requirements

- Look up powershell (this is the blue windows command line)
- It starts out in your home directory, navigate to your myanmar_xlsx_to_csv-master directory

i.e. if you downloaded and extracted it in your downloads folder then *change directory*:

```
cd C:\Users\rebeldroid12\Downloads\myanmar_xlsx_to_csv-master\myanmar_xlsx_to_csv-master
```
![windows pshell dir](https://github.com/rebeldroid12/myanmar_xlsx_to_csv/blob/master/windows-resources/in_dir_path.PNG)


*you can do `ls` to look at what is in your directory. Here you can see the main Python script which is called myanmar_xlsx_to_csv.py*


Run the requirements.txt file

```
pip install -r requirements.txt
```

![windows run req](https://github.com/rebeldroid12/myanmar_xlsx_to_csv/blob/master/windows-resources/run_requirements.PNG)

(5): Running the program. 
-Once you have opened the command line and navigated to the `myanmar_xlsx_to_csv` directory in powershell
-Locate the xlsx you would like to run and grab the path for it. 

![data xlsx](https://github.com/rebeldroid12/myanmar_xlsx_to_csv/blob/master/windows-resources/budget-data.PNG)

*i.e. `C:\Users\rebeldroid12\Downloads\data\budget-data.xlsx`. *

Go to your command line and run the Python file for that .xlsx file:

*You should be in the directory `myanmar_xlsx_to_csv-master` or know the path to the script to be able to call it*

```
python myanmar_xlsx_to_csv.py C:\Users\rebeldroid12\Downloads\data\budget-data.xlsx
```
It will print out 'All done!' when the .csv file is created. Enjoy the clean data :)

