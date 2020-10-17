#/bin/bash

# check env variables for java, WMIC, FINDSTR, and marcedit_path, otherwise things won't work.  e.g. https://kb.informatica.com/solution/23/Pages/4/156865.aspx
# assumes MarcEdit is installed
export MARCEDIT_PATH=/Applications/MarcEdit3.app/Contents/MacOS
export DateTime=$(date +"%Y%m%d_%H%M%S")
# needs Java v8 or greater
export JAVA=java
export parameters=-Xmx1024m
export CP="-cp ../vendor/saxonica/saxon9he.jar"

mkdir -p import
mkdir -p export/tmp
mkdir -p export/backup

echo Post-processing MARCXML files...

xmlfiles=(`find ./export -maxdepth 1 -name "*.xml"`)

if [ ${#xmlfiles[@]} -gt 0 ]; then
  for f in $xmlfiles; do
    fne="${f##*/}"
    filename="${fne%.xml}"
    echo $filename

    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:$f -xsl:"https://raw.githubusercontent.com/YaleArchivesSpace/ASpace-plus-MARC/master/aspace/xslt/MARCxml-post-processing.xsl" -o:export/tmp/$fne -warnings:silent

	$MARCEDIT_PATH/MarcEdit3 -s export/tmp/$fne -d export/tmp/$filename.mrc -xmlmarc

	echo "Generating import file... (please wait)"
	
	mrcfiles=(`find ./export/tmp -maxdepth 1 -name "*.mrc"`)
	#for f in $mrcfiles; do
	#	export files=$f

	$MARCEDIT_PATH/MarcEdit3 -s $mrcfiles -d import/$USER-$DateTime.mrc -join
  done

  echo "MARC file created."

  echo "Removing temporary files..."

  mv export/*.xml export/backup

  rm -r export/tmp

  echo "All done."
else
  echo "No XML filenames found. Let's stop!"
fi

read -rsp $'Press any key to continue...\n' -n1 key