Schedule Organizer
developed by Tom Zhang

Purpose:
The purpose of this program is to help school staff better visualize their class schedule.
The program will output an excel file.
See the sample output file in the sample files folder
Classes in red are full year classes.
Classes in purple are grouped classes.

In package:
The package should include class values.xlsx, departments.txt, input.txt, sample files folder, schedule.exe, and schedule.py (python source code).
Other files are neccessary dependencies and are not to be deleted.

Requirements:
The program is compatible with Windows systems.
The user must have a piece of software to open .xlsx files (most likely Microsoft Excel)
There have to be 3 files in the same directory as the executable file for the program to work: input.txt, departments.txt, and class values.xlsx.
The files have to be named exactly the same as written.
To run the program, double click on schedule.exe

********* input.txt *********
This is the input class lists to the program.
The four columns must be the course, its term, its schedule, and its department.
See the sample files folder to fetch the original file
*********

********* departments.txt *********
This file assigns each class to a department.
You may want to change this file if you have additional classes than the original list.
There are two types of exceptions to this list:
If the sixth character of the class is "F", then the program will recognize the class as a French class, and assign the class to the French department.
If the department of the class reads "ADMIN", then the program will assign the class to the Support department.

The default file contains
"""
{"CRW":0,"ELA":0,"ELB":0,"ELM":0,"ELW":0,"ARE":1,"BAN":1,"CHO":1,"CME":1,"DAN":1,"DRA":1,"GRA":1,"GUI":1,"PHO":1,"VAR":1,"VAS":1,"CAL":2,"MAF":2,"MFP":2,"MPC":2,"MST":2,"MTH":2,"MTP":2,"MWA":2,"CAR":3,"HEA":3,"PED":3,"PEF":3,"PEM":3,"WLF":3,"WLL":3,"WLM":3,"BIO":4,"CHE":4,"CSC":4,"ESC":4,"HSC":4,"PHP":4,"PHY":4,"PSC":4,"SCI":4,"HIS":5,"LAW":5,"NAT":5,"PSY":5,"SOC":5,"ACC":6,"CAC":6,"CLO":6,"CWA":6,"CWB":6,"CWE":6,"DRC":6,"ENT":6,"FLT":6,"FOO":6,"IAS":6,"IND":6,"INF":6,"INT":6,"LIF":6,"LTR":6,"PAA":6,"RBA":6,"TEC":6,"FRA":7,"FRB":7,"FRE":7,"EAA":8,"EAB":8}
"""
To add a new class to the list, put a comma before the ending curly bracket, and in double quotes, put the code of the class.
Put a colon after the double quotes, and after the colon, put a numeric value representing the department the class you want it to be in.
ELA classes have values of 0
Fine Arts classes have values of 1
Math classes have values of 2
Phys. Ed classes have values of 3
Science classes have values of 4
Social Science classes have values of 5
PAA classes have values of 6
Frech Imm. & Second Languages classes have values of 7
EAL & Support classes have values of 8

Do it exactly in this format or there will be an error.
*********

********* class values.xlsx *********
This is the file specifying the values added to the "total" column of the output excel file for each class.
By default, a value of 1 is added to the total of the corresponding grade, with 1 exception:
Classes with departments "ADMIN" and "EAL" don't have values.
You may want to change it if you want to add decimal values to the total, add values to different grade, or if you want to remove existing values.
The sheet has to be named "Sheet1"

To add a class to the file, enter the class code in the classes column, dropping the last dash character "-" and all characters behind it.
Then, enter the values you want to add for each grade.
The semester column specifies if you don't want all classes of the class type to follow the rules, and only classes of designated semesters.
enter "1" for semester 1
enter "2" for semester 2
enter "Y" for full year classes
For example, if you want full year classes and semester 1 classes to follow the rules, enter "F1" in any order of characters, then the program will not follow the rules for semester 2 classes.
*********

Upon running the program, the user should look for any outputs in the command prompt indicating any errors.