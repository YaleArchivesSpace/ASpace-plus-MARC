<?xml version="1.0" encoding="UTF-8"?>
<sch:schema xmlns:sch="http://purl.oclc.org/dsdl/schematron" queryBinding="xslt2"
    xmlns:sqf="http://www.schematron-quickfix.com/validator/process">

    <sch:ns prefix="marc" uri="http://www.loc.gov/MARC21/slim"/>

    <sch:pattern>
        <sch:rule context="marc:record">
            <sch:let name="accession" value="marc:datafield[@tag='952']/marc:subfield[@code='a']"/>
            <sch:assert test="marc:datafield[@tag='245']">
                This record, <sch:value-of select="$accession"/>, needs a 245 title to be compliant with our best practices.
            </sch:assert>
            <sch:assert test="marc:datafield[@tag='300']">
                This record, <sch:value-of select="$accession"/>, needs a 300 extent to be compliant with our best practices.
            </sch:assert>
            <sch:assert test="marc:datafield[@tag='520']">
                This record, <sch:value-of select="$accession"/>, needs a 520 abstract to be compliant with our best practices.
            </sch:assert>
            <sch:assert test="marc:datafield[@tag='524']">
                This record, <sch:value-of select="$accession"/>, needs a 524 citation to be compliant with our best practices.
            </sch:assert>
            <sch:assert test="marc:datafield[@tag='546']">
                This record, <sch:value-of select="$accession"/>, needs a 546 language note to be compliant with our best practices.
            </sch:assert>
            <sch:assert test="marc:datafield[@tag='952']">
                Hold up. This record needs an accession identifier!
            </sch:assert>
        </sch:rule>
        <sch:rule context="marc:controlfield[@tag='008']">
            <sch:assert test="matches(substring(., 8, 4), '[0-9u]{4}')">
                This 008 date 1 value is not valid.  For our best practices, it should only contain numerals or "u".
            </sch:assert>
            <sch:assert test="matches(substring(., 16, 3), '[a-z #]{3}')">
                This 008 place code is not valid.  For our best practices, it should only contain letters, a space, or a "#".
            </sch:assert>
        </sch:rule>
    </sch:pattern>

</sch:schema>
