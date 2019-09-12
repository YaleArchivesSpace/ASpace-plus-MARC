<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:math="http://www.w3.org/2005/xpath-functions/math"
    xmlns:marc="http://www.loc.gov/MARC21/slim"
    exclude-result-prefixes="xs math"
    version="3.0">
    

    <xsl:template match="@*|node()">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>
     
     <xsl:template match="marc:record">
         <xsl:variable name="rules" select="marc:datafield[@tag='040']/marc:subfield[@code='e'][1]/text()"/>
         <xsl:copy>
             <xsl:apply-templates select="@*|node()">
                 <xsl:with-param name="rules" select="$rules" tunnel="yes"/>
             </xsl:apply-templates>
         </xsl:copy>
     </xsl:template>
    
    <xsl:template match="marc:datafield[starts-with(@tag, '1')]/marc:subfield[following-sibling::marc:subfield[1][@code='e']][not(matches(normalize-space(.), '[-.:,;)/\]]$'))]">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:text>,</xsl:text>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="marc:datafield[starts-with(@tag, '1') or starts-with(@tag, '6')]/marc:subfield[following-sibling::marc:subfield[1][@code='0']][not(matches(normalize-space(.), '[-.:,;)/\]]$'))]">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:text>.</xsl:text>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="marc:datafield[@tag = '655']/marc:subfield[@code='a'][following-sibling::marc:subfield[1][@code='2']][not(matches(normalize-space(.), '[-.:,;)/\]]$'))]">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:text>.</xsl:text>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="marc:datafield[@tag = '264']/marc:subfield[following-sibling::marc:subfield[1][@code='c']][not(matches(normalize-space(.), '[-.:,;)/\]]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:if test="not($rules = 'dacs')">
                <xsl:text>,</xsl:text>
            </xsl:if>
        </xsl:copy>
    </xsl:template>
    
    <!-- always add a period at the end of 264 c except if there is already a closing bracket....  or, well, a period ] -->
    <xsl:template match="marc:datafield[@tag = '264']/marc:subfield[@code='c'][not(matches(normalize-space(.), '[.\]]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:text>.</xsl:text>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="marc:datafield[@tag = '561']/marc:subfield[last()][not(matches(normalize-space(.), '[.]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:text>.</xsl:text>
        </xsl:copy>
    </xsl:template>
    
    <!-- 300 $a  : $b ; $c OR $a ; $c -->
    <xsl:template match="marc:datafield[@tag = '300']/marc:subfield[@code='a'][following-sibling::marc:subfield[@code='b']][1][not(matches(normalize-space(.), '[:]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:if test="not($rules = 'dacs')">
                <xsl:text> :</xsl:text>
            </xsl:if>
        </xsl:copy>
    </xsl:template>
    <xsl:template match="marc:datafield[@tag = '300']/marc:subfield[@code='b'][following-sibling::marc:subfield[@code='c']][1][not(matches(normalize-space(.), '[:]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:if test="not($rules = 'dacs')">
                <xsl:text> :</xsl:text>
            </xsl:if>
        </xsl:copy>
    </xsl:template>
    <xsl:template match="marc:datafield[@tag = '300']/marc:subfield[@code='a'][following-sibling::marc:subfield[@code='c']][1][not(matches(normalize-space(.), '[;]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:if test="not($rules = 'dacs')">
                <xsl:text> ;</xsl:text>
            </xsl:if>
        </xsl:copy>
    </xsl:template>
    
</xsl:stylesheet>