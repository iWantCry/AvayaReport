#Avaya Call Volume Report Program
#User Manual


1.1	Purpose of the program:
•	The purpose of this program is to generate the daily Avaya call volume report for the day before. This program will semi-automate most steps taken to generate the final report.

1.2	Description of the program:
•	This program will consist of 2 parts:
o	Generating the Avaya report from the raw data
o	Generate the final report for the email

1.3	Important requirements before running the program:
•	Only run this program on your local computer (Not OneDrive) or the program will not work as intended. If the user does end up running the program to edit files in OneDrive, the program will warn the user of such an error.
•	Export all raw files from Avaya IPOCC before moving on to run the program. All of the raw files must be in the same directory.














2.1	How to use the program (GenerateAvayaReport):
•	After exporting the raw files, run the executable file for “GenerateAvayaReport”. This will generate the Avaya report (Avaya Report file). The program will still make use of the previous VBA macro to create the report.

•	After running the program, it will ask the user for the directory of all the exported raw data files from Avaya IPOCC.


•	Select the directory and the program will conduct a check to see if all the required files are included. This includes the template file, macro file, previous day’s report file and the CRV file. In total there should be 8 files in this directory.

•	After a successful check, the program will generate the report file.


•	After the program has finished running, the user has to manually change the values on the excel pivot table slicers and copy the table data on to the avaya template file.

2.2 How to use the program (GenerateEmailReport):
•	After pasting the table data into the template file, save and close the avaya file as the program will not be able to run if the avaya file is open as excel will see it as 2 separate authors (The user and the program) trying to edit the same file.

•	In the case that the user accidentally leaves the file open, the program will stop and warn the user of the error.

•	 After ensuring the previous steps were done correctly, run the executable file “GenerateEmailReport”.

•	The program will ask the user for the location of the avaya template file so the formatting can be completed.

•	After giving the program the location of the template file, it will automatically format the table.




2.3 After using the program:
•	The completed and finalised report will be found in the avaya template file.
•	Since the program has been running on your local computer, the files need to be manually moved into OneDrive and SQM.

