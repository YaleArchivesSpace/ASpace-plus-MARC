#/bin/bash

# assuming that marcedit this has bene installed, like so
export MARCEDIT_PATH=/Applications/MarcEdit3.app/Contents/MacOS
# still assuming this is alredy set / available
export JAVA=java
# $MARCEDIT_PATH/MarcEdit3
export parameters=-Xmx1024m
export CP="-cp ../vendor/saxonica/saxon9he.jar"

export DateTime=$(date +"%F-%T")

mkdir -p import
mkdir -p spreadsheets/tmp
mkdir -p spreadsheets/backup

xmlfiles=(`find ./spreadsheets -maxdepth 1 -name "*.xml"`)
if [ ${#xmlfiles[@]} -gt 0 ]; then
  for f in ${xmlfiles[@]}; do
    fne="${f##*/}"
    filename="${fne%.xml}"
    echo $filename

    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:$f -xsl:"xslt/ExcelXML-to-MarcXML.xsl" -o:spreadsheets/tmp/$fne -warnings:silent
    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:spreadsheets/tmp/$fne -xsl:"xslt/MarcXML-add-6xx.xsl" -o:spreadsheets/tmp/$fne -warnings:silent
    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:spreadsheets/tmp/$fne -xsl:"xslt/MarcXML-reorder-and-prep.xsl" -o:spreadsheets/tmp/$fne -warnings:silent
    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:spreadsheets/tmp/$fne -xsl:"xslt/Deal-with-ISBD-issues.xsl" -o:spreadsheets/tmp/$fne -warnings:silent

    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:xslt/marc-tests.sch -xsl:"../vendor/schematron/iso_dsdl_include.xsl" -o:schematron/marc-tests-1.sch
    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:schematron/marc-tests-1.sch -xsl:"../vendor/schematron/iso_abstract_expand.xsl" -o:schematron/marc-tests-2.sch
    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:schematron/marc-tests-2.sch -xsl:"../vendor/schematron/iso_svrl_for_xslt2.xsl" -o:schematron/schematron.xsl
    $JAVA $parameters $CP net.sf.saxon.Transform -t -s:spreadsheets/tmp/$fne -xsl:"schematron/schematron.xsl" -o:reports/${DateTime}_error_report_${filename}.svrl

    $MARCEDIT_PATH/MarcEdit3 -s spreadsheets/tmp/$fne -d import/$filename.mrc -xmlmarc
  done

  echo "MARC file created."

  echo "Moving input files to backup directory..."

  mv spreadsheets/*.xml spreadsheets/backup

  rm -r spreadsheets/tmp schematron

  echo "All done."

else
  echo "No XML filenames found. Let's stop!"
fi


read -rsp $'Press any key to continue...\n' -n1 key
