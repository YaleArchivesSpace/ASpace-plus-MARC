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
  for f in ${xmlfiles[@]}; do
    fne="${f##*/}"
    filename="${fne%.xml}"
    echo $filename

    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:$f -xsl:"https://raw.githubusercontent.com/YaleArchivesSpace/ASpace-plus-MARC/master/aspace/xslt/MARCxml-post-processing.xsl" -o:export/tmp/$fne -warnings:silent
    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:export/tmp/$fne -xsl:"https://raw.githubusercontent.com/YaleArchivesSpace/ASpace-plus-MARC/master/excel/xslt/MarcXML-reorder-and-prep.xsl" -o:export/tmp/$fne -warnings:silent
    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:export/tmp/$fne -xsl:"https://raw.githubusercontent.com/YaleArchivesSpace/ASpace-plus-MARC/master/excel/xslt/Deal-with-ISBD-issues.xsl" -o:export/tmp/$fne.xml -warnings:silent

    $MARCEDIT_PATH/MarcEdit3 -s export/tmp/$fne -d import/$filename.mrc -xmlmarc

  done

  echo "Generating import file... (please wait)"

  mrcfiles=(`find export/tmp -maxdepth 1 -name "*.mrc"`)

  # MarcEdit requirement is to send the list of files as a semi-colon separated list, so i'd think this should work, but alas it does not.
  files=$(IFS=\;; echo "${mrcfiles[*]}")

  echo $files

  # Since I cannot get the join option to work with MarcEdit 3, the way it works on Windows, I'm giving up on that for now.... and moving those XML->MRC files over to the import directory instead.
  # $MARCEDIT_PATH/MarcEdit3 -s $files -d import/$USER-$DateTime.mrc -join

  # echo "MARC file created."

  echo "Removing temporary files..."

  mv export/*.xml export/backup

  rm -r export/tmp

  echo "All done."
else
  echo "No XML filenames found. Let's stop!"
fi

read -rsp $'Press any key to continue...\n' -n1 key
