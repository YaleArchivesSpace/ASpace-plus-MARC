rem check env variables for java, WMIC, FINDSTR, and marcedit_path, otherwise things won't work.  e.g. https://kb.informatica.com/solution/23/Pages/4/156865.aspx
@echo off
FOR /f %%a IN ('WMIC OS GET LocalDateTime ^| FIND "."') DO SET DTS=%%a
SET DateTime=%DTS:~0,8%_%DTS:~8,6%
SET JAVA=java
SET parameters=-Xmx1024m
SET CP=-cp ..\vendor\saxonica\saxon9he.jar
setlocal enabledelayedexpansion

if not exist import mkdir import
if not exist export\tmp mkdir export\tmp
if not exist export\backup mkdir export\backup

echo Post-processing MARCXML files...

 for %%f in (export\*.xml) do (
            echo %%~nf
			
	%JAVA% %parameters% %CP% net.sf.saxon.Transform -t -s:export\%%~nf.xml -xsl:"https://raw.githubusercontent.com/YaleArchivesSpace/ASpace-plus-MARC/master/aspace/xslt/MARCxml-post-processing.xsl" -o:export\tmp\%%~nf.xml -warnings:silent

	"%marcedit_path%\cmarcedit.exe" -s export\tmp\%%~nf.xml -d export\tmp\%%~nf.mrc -xmlmarc

    )

echo Generating import file... (please wait)
	
 for %%n in ("export\tmp\*.mrc") DO ( 
	SET files=!files!%%n;
 )

"%marcedit_path%\cmarcedit.exe" -s %files% -d import\%USERNAME%-%DateTime%.mrc -join

echo MARC file created.

echo Removing temporary files...


move export\*.xml export\backup


rmdir export\tmp /S /Q

echo All done. 

pause