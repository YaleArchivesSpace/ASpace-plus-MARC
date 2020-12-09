# Processing single files exported from ArchivesSpace from MarcXML to MARC

# Current process at https://docs.google.com/document/d/189W6dSrjhmmDly6woaxEKZCwVZXhmpIVbFnlCcuEkE4/edit

# Yet More:
# See also https://guides.library.yale.edu/c.php?g=296249&p=7060320
# 	https://www.youtube.com/watch?v=2KiDd_nw884
# And https://docs.google.com/document/d/1DI_7YNZy-RcjQ9hpMMbxJEkHFpYndzmDoG3ylOc38BY/edit#heading=h.bb4v2dv2z11q ("Exporting Records" section of local user manual)
# AND
# https://drive.google.com/open?id=1q4x_EQ8zExqeD4NIEgvS11GuFD_nALvr (slide show from a 2016[?] preso on the matter)


# This file is intended to supersede ASpace

# from pymarc import MARCReader
from pymarc import parse_xml_to_array
from pymarc import MARCReader
from datetime import datetime
from pathlib import Path

import getpass
import glob
import subprocess

# Establish datetime string as %Y%m%d_%H%M%S
# get current date and time
now = datetime.now()
# Initialize format string
format = "%Y%m%d_%H%M%S"
# Format datetime variable
DateTime = now.strftime(format)
 
# Initialize variables for Java name, Java memory, cp flag
java_name = "java"
parameters = "-Xmx1024m"
CP = "../vendor/saxonica/saxon9he.jar"

# Make directory for `import` if it doesn't exist
# https://stackoverflow.com/a/273227
Path("./import").mkdir(parents=True, exist_ok=True)
# Make subdirectories under `export` if they don't exist
Path("./export/tmp").mkdir(parents=True, exist_ok=True)
Path("./export/backup").mkdir(parents=True, exist_ok=True)

# Message user what we are doing
print("Post-processing MARCXML files...")

# Find all XML files directly under `export`
mylist = [f for f in glob.glob("./export/*.xml")]
# If there are any files
if (len(mylist) > 0):
	# Loop through them
	for filepath in mylist:
		# Initialize variable for file's current file name? [why?]
		oldfile = Path(filepath).name
		# Initialize variable for file name of current name less extension
		# We know the extension is .xml so we can hard code the amount to trim
		# https://stackoverflow.com/a/663175
		newfile = oldfile[:-4]
		# Run Java (Saxon) XSLT
		# Thank you to https://www.datasciencelearner.com/how-to-call-jar-file-using-python/
		# If the calls to GitHub won't get the remote XSL, replace with local version
		# '-xsl:./xslt/MARCxml-post-processing.xsl'
		# May need to double quote the XSL file string
		subprocess.call([java_name, parameters, '-cp', CP, 'net.sf.saxon.Transform', '-t', '-s:export/'+oldfile,  '-xsl:https://raw.githubusercontent.com/YaleArchivesSpace/ASpace-plus-MARC/master/aspace/xslt/MARCxml-post-processing.xsl', '-o:export/tmp/'+oldfile, '-warnings:silent'])
		subprocess.call([java_name, parameters, '-cp', CP, 'net.sf.saxon.Transform', '-t', '-s:export/tmp/'+oldfile,  '-xsl:https://raw.githubusercontent.com/YaleArchivesSpace/ASpace-plus-MARC/master/excel/xslt/MarcXML-reorder-and-prep.xsl', '-o:export/tmp/'+oldfile, '-warnings:silent'])
		subprocess.call([java_name, parameters, '-cp', CP, 'net.sf.saxon.Transform', '-t', '-s:export/tmp/'+oldfile,  '-xsl:https://raw.githubusercontent.com/YaleArchivesSpace/ASpace-plus-MARC/master/excel/xslt/Deal-with-ISBD-issues.xsl', '-o:export/tmp/'+oldfile, '-warnings:silent'])

		# Convert new XML file to MARC
		records = parse_xml_to_array('export/tmp/'+oldfile)
		for record in records:
			with open('export/tmp/'+newfile+'.mrc', 'wb') as out:
				out.write(record.as_marc())
		print('Generating import file . . . (please wait)')
		# Concatenate MARC files into one
		my_marcs = [f for f in glob.glob("./export/tmp/*.mrc")]
		for marcfile in my_marcs:
			with open(marcfile, 'rb') as fh:
				reader = MARCReader(fh)
				for record in reader:
					with open(getpass.getuser()+"-"+DateTime+".mrc", ab) as out:
						out.write(record.as_marc())
		# Move file to backup directory
		x = Path(filepath)
		target = Path('export/backup/'+oldfile)
		x.rename()

	print("MARC file created.")

print("Removing temporary files . . .")
for f in Path('export/tmp').glob('*.xml'):
	f.unlink()
for f in Path('export/tmp').glob('*.mrc'):
	f.unlink()
print("Removing temporary directory . . .")
Path('export/tmp').rmdir

print("All done.")

# End looping through XML files



# ----
# >>> with open('./excel/import/spicher_test.mrc', 'rb') as fh:
# ...     reader = MARCReader(fh)
# ...     for record in reader:
# ...             print(record.title())
