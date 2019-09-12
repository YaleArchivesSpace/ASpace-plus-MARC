# ASpace-plus-MARC

## What's this all about?

First of all, everything here is still a work in progress! But the general idea behind this is that staff have identified two different use cases for repurprosing data from ArchivesSpace and moving that description into our ILS, Voyager. 

1. The first use case is to export a MARC record from an ArchivesSpace Resource record and then import that record into Voyager so that an archivist does not have to perform any rekeying.
2. The second use case is to extract data from ArchivesSpace into a spreadsheet, which can then be used by an archivist to create multiple MARC records all in one go that could then be used to import multiple records into Voyager.

Right now, both of these use cases share the same requirements, although in the future we plan to further abstract and automate both workflows.  For now, the requirements are:

* Windows, since both workflows use simple batch commands to control the transformation steps.
* Java 8 (or higher), since both workflows utilize SaxonHE to run the XML transformations.
* MarcEdit, since both workflows run MarcEdit from the commandline to convert (and combine) MARCXML records into binary MARC files.

For more information about the current setup process, see https://docs.google.com/document/d/189W6dSrjhmmDly6woaxEKZCwVZXhmpIVbFnlCcuEkE4/edit?usp=sharing 

