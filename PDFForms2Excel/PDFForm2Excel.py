#!/usr/bin/env python
# coding: utf-8
# Copyright 2017-2019 by Benedikt Sürig
# Python Script to read PDF Form Field content from a folder or subfolder Structure (-R)
# and write it to an Excel file with one Line per PDF File or one column per PDF File (-c)
# Version 2019-11-16
from os import listdir, walk
from os.path import isfile, join, abspath
from sys import exit
from PyPDF2 import PdfFileReader
import pandas as pd
import argparse 


def PDFForm2Excel(mypath, outfile):
	errcount=0
	mypath = abspath(mypath)
	pdffiles=[]
	#Put PDF Files into Array pdffiles
	if parser.parse_args().boolean_recursive == False:   #non recursive: Only input Folder is analyzed
		pdffiles = sorted( [ join(mypath, f) for f in listdir(mypath) if f[-4:]=='.pdf' and  isfile(join(mypath, f)) ])
	if parser.parse_args().boolean_recursive == True: #recursive: input Folder and Subfolders are analyzed
		for r, d, f in walk(mypath):
			for file in f:
				if  file[-4:]=='.pdf' in file:
					pdffiles.append(join(abspath(r), file))


	#Start with Initial PDF to create Dataframe schema 
	#It needs to be assured that the first PDF and its Form Fields are always readable
	#They create the master schema for all following PDFs
	#It also has to be assured that the follwing PDF Forms have the same fields in the same order
	#print(pdffiles)
	pdf = pdffiles[0]
	# pdf
	#Define Objects f (file) and fields 
	f = PdfFileReader(pdf)
	fields = f.getFields()
	#Define Dataframe Object for Results 
	results = pd.DataFrame()
	#Define dataframe for Field Name List
	try:
		df = pd.DataFrame( [(k,k1,v1) for k,v in fields.items()  for k1,v1  in v.items()], columns = ['Field','Type',pdf])
		df2 = df.loc[df['Type'] == '/V' ]       # Filter for Values := '/V' only
		df2 = df2.filter(items=['Field',pdf])
		df2 = df2.reset_index(drop = True)  # Reset Row Index
		results = df2.set_index('Field')
		results = results.drop(pdf, axis='columns')
	except:
			exit("##initial pdf read error## " + pdf)
		#Loop through all PDF Files in Input Folder
	for pdf in pdffiles:
		f = PdfFileReader(pdf) # PDF Fileobject
		producer = f.getDocumentInfo().producer
		try:
			fields = f.getFields()
			df = pd.DataFrame( [(k,k1,v1) for k,v in fields.items()  for k1,v1  in v.items()], columns = ['Field','Type',pdf])
			df2 = df.loc[df['Type'] == '/V' ]      # Filter for Values := '/V' only
			df2 = df2.filter(items=['Field',pdf])  # Add Filename as ColumnHeader 
			df2 = df2.set_index('Field') # Set 'Field' as Row Index
			df2[pdf]=df2[pdf].map(lambda x: x.lstrip('='))   #remove heading '=' from future excel Cells
			results = results.merge(df2, on='Field', how='left')  # Write Values to Array 'results'
			print("read success   " + str(producer) + " " + pdf)
		except:
			results[pdf] = '' 			#Create empty Column in results for failed PDF
			print("##read error## " + str(producer) + " " + pdf)
			errcount=errcount +1
			continue
	print('--------------')
	print( "Summary: " +str(len(pdffiles)) +" files read with "+ str(errcount) + " file read Errors")
	#write results to Excel
	try:
		if  parser.parse_args().boolean_col==False:
		 	results.T.to_excel(mypath + outfile, header=True, index=True) # Write Result dataframe in one line per PDF to Excel File
		else:
			results.to_excel(mypath + outfile, header=True, index=True)  # Write Result dataframe in one Coumn per PDF to Excel File
		print("File " + mypath + outfile + " successfully written")
	except:
		print("Error writing File " + mypath + outfile )


parser = argparse.ArgumentParser(description="PDFForm2Excel.py reads PDF Form fields from all PDFs in input folder and generates an output file in XLSX format based on the form fieldlist of the first read PDF. Benedikt Suerig 2019")
parser.add_argument('-c', '--transpose', action='store_true', default=False, dest='boolean_col', help='Output in one Column per PDF File (default: one line per PDF File)')
parser.add_argument('-R', '--recursive', action='store_true', default=True, dest='boolean_recursive', help='Read PDF Files recursive also in Subfolders of Input folder(default: non recursive)')
parser.add_argument('-i', '--input', type=str, default= '.\\Forms\\', help='input path (default: .\Forms\)')
parser.add_argument('-o', '--output', type=str, default='PDFForm2Excel.xlsx', help="output file (default: 'PDFForm2Excel.xlsx')")
args=parser.parse_args()
print(args)
print("input folder: " + args.input)
print("output file: " + args.output)
if args.boolean_recursive== True:
	print("read folders recursively")
if args.boolean_col == True:
	print("transpose output (one column per PDF")
print('--------------')
PDFForm2Excel(args.input,args.output)
