<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:math="http://www.w3.org/2005/xpath-functions/math"
    xmlns:marc="http://www.loc.gov/MARC21/slim"
    exclude-result-prefixes="xs math"
    version="3.0">
    
    <xsl:template match="@*|node()" mode="#all">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>
    
    <!-- could handle this in the first transformation, but if it's always going to be a straight up copy, it'll likely be more succinct to just copy it after the fact like so. -->
    <xsl:template match="marc:datafield[starts-with(@tag, '1')]" priority="2">
        <xsl:copy-of select="."/>
        <xsl:variable name="tag" select="string(xs:integer(@tag) + 500)"/>
        <marc:datafield tag="{$tag}">
            <xsl:attribute name="ind2">
                <xsl:value-of select="0"/>
            </xsl:attribute>
            <xsl:apply-templates select="@ind1|node()" mode="copy-6xx"/>
        </marc:datafield> 
    </xsl:template>
    
    <xsl:template match="marc:subfield[@code='e']" mode="copy-6xx"/>
    
</xsl:stylesheet>