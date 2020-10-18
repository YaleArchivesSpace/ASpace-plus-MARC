<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:math="http://www.w3.org/2005/xpath-functions/math"
    xmlns:marc="http://www.loc.gov/MARC21/slim"
    exclude-result-prefixes="xs math"
    version="3.0">
    
    <!-- still to check (maybe in the schematron?):
        also, lots of 300 examples end with punctuation.... 
        but in our case, we'll only allow a closed parantheses to end the 300.
    
The following fields never end in punctuation: 040, 043, 035, 300, 336, 337, 338
-->
   
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
    
    <!-- the following fields must end with a terminal period.  add one if none present.-->
    <xsl:template match="marc:datafield[@tag = ('245', '351', '500', '506', '561', '546',  '524')]/marc:subfield[last()][not(matches(normalize-space(.), '[.]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:text>.</xsl:text>
        </xsl:copy>
    </xsl:template>
    
    <!-- the following fields end with some sort of terminal punctuation.  add a period if no other terminal punctuation is present.-->
    <!-- regex should match:
        [!"\#$%&'()*+,\-./:;<=>?@\[\\\]^_`{|}~] 
        Change if that's too liberal.
    -->
    <xsl:template match="marc:datafield[@tag = ('520', '545')]/marc:subfield[last()][not(matches(normalize-space(.), '[\p{P}]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:text>.</xsl:text>
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
    
    <!-- 
    The following fields have varying punctuation:
    100, 600, 610, 611, 630, 650, 651, 700: Ends with a period, unless it ends with a closing parenthesis (e.g. Griggs (Family)) or a dash (open date e.g. 1821-).
    -->
    <xsl:template match="marc:datafield[@tag = ('100', '600', '610', '611', '630', '650', '651', '700')]/marc:subfield[not(@code=('0','1'))][last()][not(matches(normalize-space(.), '[-.)]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:text>.</xsl:text>
        </xsl:copy>
    </xsl:template>
    
    <!--
        If 245 $f is present, it is always preceded by a comma. 
        -->
    <xsl:template match="marc:datafield[@tag = '245']/marc:subfield[following-sibling::marc:subfield[1][@code='f']][not(matches(normalize-space(.), '[,]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
                <xsl:text>,</xsl:text> 
        </xsl:copy>
    </xsl:template>
    
    <!-- if 245 $ k is present, it should also be preceded with ' : ' -->
    <xsl:template match="marc:datafield[@tag = '245']/marc:subfield[following-sibling::marc:subfield[@code='k']][1][not(matches(normalize-space(.), '[:]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:text> : </xsl:text> 
        </xsl:copy>
    </xsl:template>
    
    <!--
        There is always a comma between 264 $a and 264 $c 
        -->
    <xsl:template match="marc:datafield[@tag = '264']/marc:subfield[@code = 'a'][following-sibling::marc:subfield[1][@code='c']][not(matches(normalize-space(.), '[,]$'))]">
        <xsl:param name="rules"/>
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:text>,</xsl:text>
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
    
    <!--
    The following fields have a period or closing parenthesis between $a and $2: 655, 656
    if neither are present, add a period.
    -->
    <xsl:template match="marc:datafield[@tag = ('655', '656')]/marc:subfield[@code='a'][following-sibling::marc:subfield[1][@code='2']][not(matches(normalize-space(.), '[.)]$'))]">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
            <xsl:text>.</xsl:text>
        </xsl:copy>
    </xsl:template>
    
</xsl:stylesheet>