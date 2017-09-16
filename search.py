#!/usr/bin/env python

import re
import xlsxwriter
from datetime import datetime
import dateutil
from dateutil.relativedelta import *

workbook = xlsxwriter.Workbook('cases_new.xlsx')
worksheet = workbook.add_worksheet()
outputRow = 0

formatDate = workbook.add_format()
formatDate.set_num_format('dd/mm/yyyy')

formatTitle = workbook.add_format({'bg_color':'#FFA0A0'})

worksheet.write_string(outputRow,0,"IPP",formatTitle)
worksheet.write_string(outputRow,1,"sampleID",formatTitle)
worksheet.write_string(outputRow,2,"name",formatTitle)
worksheet.write_string(outputRow,3,"gender",formatTitle)
worksheet.write_string(outputRow,4,"birthdate",formatTitle)
worksheet.write_string(outputRow,5,"sample date",formatTitle)
worksheet.write_string(outputRow,6,"age at sampling",formatTitle)
worksheet.write_string(outputRow,7,"clinical info",formatTitle)
worksheet.write_string(outputRow,8,"diagnostic",formatTitle)

worksheet.freeze_panes(1, 0)

with open("TESTDATA.txt") as target:
#with open("ALLDATA.txt") as target:
    for line in target:
        match = re.match("Rapport anatomo-pathologique\s+Examen N°\s+(?P<sampleID>H\d{7})", line)
        if not match:
            continue
        sampleID = match.group('sampleID')

        match = False
        while not match:
            line = next(target)
            match = re.match("Patient\s+(?P<name>[A-Z ]+,[A-Z ]+)\s+\((?P<gender>[FM])\)\s+Date de prélèvement :\s+(?P<sampleDate>\d{1,2}\.\d{1,2}\.\d{4})", line)
        name = match.group('name')
        gender = match.group('gender')
        sampleDate = datetime.strptime(match.group('sampleDate'),'%d.%m.%Y')    

        match = False
        while not match:
            line = next(target)
            match = re.match("né\(e\)\s+le\s+(?P<birthDate>\d{1,2}\.\d{1,2}\.\d{4})", line)
        birthDate = datetime.strptime(match.group('birthDate'),'%d.%m.%Y') 

        ageAtSampling = dateutil.relativedelta.relativedelta(sampleDate, birthDate).years

        outputRow+=1
        worksheet.write_string(outputRow,1,sampleID)
        worksheet.write_string(outputRow,2,name)
        worksheet.write_string(outputRow,3,gender)
        worksheet.write_datetime(outputRow,4,birthDate,formatDate)
        worksheet.write_datetime(outputRow,5,sampleDate,formatDate)
        worksheet.write_number(outputRow,6,ageAtSampling)
        
        match = False
        while not match:
            line = next(target)
            if re.match("tél: 0", line):
                print("Broken record for name: %s" % name)
                break                
            match = re.match("Diagnostic :", line)
        if not match:
            continue
        lineList = ""
        nextLine = next(target)
        while nextLine != "\n":
            lineList += nextLine
            nextLine = next(target)
        diagnostic = lineList

        match = False
        while not match:
            line = next(target)
            if re.match("tél: 0", line):
                print("Broken record for name: %s" % name)
                break                
            match = re.match("Renseignements cliniques :", line)
        if not match:
            continue
        lineList = ""
        nextLine = next(target)
        while nextLine != "\n":
            lineList += nextLine
            nextLine = next(target)
        clinical = lineList

        worksheet.write_string(outputRow,7,clinical)
        worksheet.write_string(outputRow,8,diagnostic)

        print("name:%s"%name)

worksheet.autofilter(0,0,outputRow,8)

workbook.close()
target.closed

