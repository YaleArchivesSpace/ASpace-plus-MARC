rem check env variables for java, WMIC, FINDSTR, and marcedit_path, otherwise things won't work.  e.g. https://kb.informatica.com/solution/23/Pages/4/156865.aspx
@echo off

SET JAVA=java
SET parameters=-Xmx1024m
SET CP=-cp ..\vendor\saxonica\saxon9he.jar

setlocal enabledelayedexpansion

if not exist import mkdir import
if not exist spreadsheets\tmp mkdir spreadsheets\tmp
if not exist spreadsheets\backup mkdir spreadsheets\backup

 for %%f in (spreadsheets\*.xml) do (
            echo %%~nf
			
	%JAVA% %parameters% %CP% net.sf.saxon.Transform -t -s:spreadsheets\%%~nf.xml -xsl:"xslt\ExcelXML-to-MarcXML.xsl" -o:spreadsheets\tmp\%%~nf.xml -warnings:silent

	%JAVA% %parameters% %CP% net.sf.saxon.Transform -t -s:spreadsheets\tmp\%%~nf.xml -xsl:"xslt\MarcXML-add-6xx.xsl" -o:spreadsheets\tmp\%%~nf.xml -warnings:silent

	%JAVA% %parameters% %CP% net.sf.saxon.Transform -t -s:spreadsheets\tmp\%%~nf.xml -xsl:"xslt\MarcXML-reorder-and-prep.xsl" -o:spreadsheets\tmp\%%~nf.xml -warnings:silent
	
	%JAVA% %parameters% %CP% net.sf.saxon.Transform -t -s:spreadsheets\tmp\%%~nf.xml -xsl:"xslt\Deal-with-ISBD-issues.xsl" -o:spreadsheets\tmp\%%~nf.xml -warnings:silent
	
	
	"%marcedit_path%\cmarcedit.exe" -s spreadsheets\tmp\%%~nf.xml -d import\%%~nf.mrc -xmlmarc

    )

echo MARC file created.

echo Moving input files to backup directory...

move spreadsheets\*.xml spreadsheets\backup

rmdir spreadsheets\tmp /S /Q

echo All done. 

pause