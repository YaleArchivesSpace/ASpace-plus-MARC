<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:math="http://www.w3.org/2005/xpath-functions/math"
    xmlns:marc="http://www.loc.gov/MARC21/slim" xmlns:mdc="http://mdc"
    exclude-result-prefixes="xs math" version="3.0">
    <xsl:output method="xml" indent="true" encoding="UTF-8"/>

    <xsl:variable name="sortOrder">
        <mdc:sort>
            <mdc:item tag="506"/>
            <mdc:item tag="540"/>
            <mdc:item tag="545"/>
            <mdc:item tag="520"/>
            <mdc:item tag="580"/>
            <mdc:item tag="555"/>
            <mdc:item tag="505"/>
            <mdc:item tag="510"/>
            <mdc:item tag="581"/>
            <mdc:item tag="590"/>
            <mdc:item tag="530"/>
            <mdc:item tag="533"/>
            <mdc:item tag="535"/>
            <mdc:item tag="561"/>
            <mdc:item tag="546"/>
            <mdc:item tag="562"/>
            <mdc:item tag="500"/>
            <mdc:item tag="502"/>
            <mdc:item tag="544"/>
            <mdc:item tag="524"/>
        </mdc:sort>
    </xsl:variable>

    <xsl:template match="@* | node()">
        <xsl:copy>
            <xsl:apply-templates select="@* | node()"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="marc:subfield">
        <xsl:copy>
            <xsl:apply-templates select="@*"/>
            <!-- here's where we chomp all leading and trailing whitespace, etc. -->
            <xsl:value-of select="normalize-space(.)"/>
        </xsl:copy>
    </xsl:template>

    <xsl:template match="marc:record">
        <xsl:copy>
            <xsl:apply-templates select="@* | marc:leader | marc:controlfield"/>
            <xsl:apply-templates select="marc:datafield">
                <!-- here's our wacky sort algorithm:
                    we're assuming all control fields are already  in order.  for datafields, if the tag is less than 500 or greather than 599
                    then we'll use the tag as the sort order.
                    if the tag is a 5xx tag, then we use the $sortOrder variable above to put things in order.
                    506 winds up as 500
                    540 winds up as 501, etc.
                    -->
                <xsl:sort select="if (xs:integer(current()/@tag) lt 500 or xs:integer(current()/@tag) gt 599) then @tag 
                    else count($sortOrder/mdc:sort/mdc:item[@tag eq current()/@tag]/preceding-sibling::*) + 500" data-type="number"/>
            </xsl:apply-templates>
        </xsl:copy>
    </xsl:template>
    
    <!-- any other values to add defaults for is missing? -->
    <xsl:template match="marc:datafield[@tag='655'][not(marc:subfield/@code='2')]">
        <xsl:variable name="source" select="if (lower-case(.) = 'manuscript maps') then 'lcgft' else 'aat'"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <marc:subfield code='2'>
                <xsl:value-of select="$source"/>
            </marc:subfield>
        </xsl:copy>
    </xsl:template>

</xsl:stylesheet>
