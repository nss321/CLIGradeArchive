# CLIGradeArchive
Simple Grading System based on CLI

Using Python, SQLite3 and openpyxl

You can manage some 'Personal Grade data'. It runs on Python terminal.

It has Basic CRUD func, and you can export raw data from database to .xlsx file.

Database schema 

CREATE TABLE "student" ( "Student_Name" TEXT UNIQUE, "Student_ID" TEXT UNIQUE, "Score_Kor" INTEGER, "Score_Eng" INTEGER, "Score_Math" INTEGER, 'Score_Avg' INTEGER, UNIQUE("Student_ID") )

