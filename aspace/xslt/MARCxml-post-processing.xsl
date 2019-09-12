<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:map="http://www.w3.org/2005/xpath-functions/map"
    xmlns:marc="http://www.loc.gov/MARC21/slim" xmlns:fn="http://www.w3.org/2005/xpath-functions"
    exclude-result-prefixes="xs marc fn map" version="3.0">
    
    <xsl:output method="xml" encoding="UTF-8" indent="yes"/>
    
    <!-- to do:
        parse stowaway subfields from subfield Q and put in their right place.  example:
        
    <datafield ind1="2" ind2="0" tag="610">
      <subfield code="a">Olympic Games,</subfield>
      <subfield code="b"> </subfield>
      <subfield code="e">subject,</subfield>
      <subfield code="g">($d: 1908 $c: London, England),</subfield>
      <subfield code="n">(4th).</subfield>
      <subfield code="0">http://id.loc.gov/authorities/names/n85302725</subfield>
    </datafield>
    
    and yes, that 610 should be a 611.  We're thinking of another hack for that.  One hack leads to another!
    
    <marc:datafield xmlns:marc="http://www.loc.gov/MARC21/slim" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" tag="611" ind1="2" ind2="0">
      <marc:subfield code="a">Olympic Games</marc:subfield>
      <marc:subfield code="n">(4th :</marc:subfield>
      <marc:subfield code="d">1908 :</marc:subfield>
      <marc:subfield code="c">London, England)</marc:subfield>
      <marc:subfield code="0">http://id.loc.gov/authorities/names/n85302725</marc:subfield>
    </marc:datafield>
    
        -->
    
    <!-- we could add a map of all values and select a default... but since we always want 'i' right now, that's what we'll provide -->
    <xsl:param name="descriptive-cataloging-form" select="'i'"/>
    
    <xsl:param name="place-of-publication" select="'ctu'"/>
    
    <xsl:param name="default-language-of-cataloging" select="'eng'"/>
    
    <xsl:variable name="eadid" select="marc:collection/marc:record[1]/marc:datafield[@tag eq '856']/marc:subfield[@code eq 'u']/substring-after(., 'http://hdl.handle.net/10079/fa/')"/>
    
    <!-- as usual, the standard identity template -->
    <xsl:template match="@*|node()">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>
    
    <!-- shouldn't be necessary, but ASpace puts the namespace after the schema file within xsi:schemaLocation.
        so here, until we override that-->
    <xsl:template match="marc:collection">
        <marc:collection xmlns="http://www.loc.gov/MARC21/slim" xmlns:marc="http://www.loc.gov/MARC21/slim"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
            xsi:schemaLocation="http://www.loc.gov/MARC21/slim http://www.loc.gov/standards/marcxml/schema/MARC21slim.xsd">
            <xsl:apply-templates select="node()"/>
        </marc:collection>
    </xsl:template>
    
    <!-- please note that MARC counts characters starting at 0, whereas XPath starts at 1.
        because of that, we are replacing the 18th character in the leader field by going after 
        the single character in the 19th position of the substring in the select statement below-->
    <xsl:template match="marc:leader">
        <xsl:copy>
            <!-- although we could also apply templates for comments and processing instructions here, those are never going to be in the source file, so we'll just worry about attributes -->
            <xsl:apply-templates select="@*"/>
            <xsl:value-of select="substring(., 1, 18) || $descriptive-cataloging-form || substring(., 20)"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="marc:controlfield[@tag eq '008']">
        <xsl:copy>
            <xsl:apply-templates select="@*"/>
            <xsl:value-of select="substring(., 1, 15) || $place-of-publication || substring(., 19)"/>
        </xsl:copy>
        <xsl:if test="$eadid">
            <datafield ind1=" " ind2=" " tag="035">
                <subfield code="9">
                    <xsl:value-of select="'YUL(ead).' || $eadid"/>
                </subfield>
            </datafield>
        </xsl:if>
    </xsl:template>
    
    <!-- should update the repo records in ASpace not to include "US-" in the agency code, but we'll also scrub this after the export process -->
    <xsl:template match="marc:datafield[@tag eq '040']/marc:subfield[@code = ('a', 'c')][starts-with(., 'US-')]">
        <xsl:copy>
            <xsl:apply-templates select="@*"/>
            <xsl:value-of select="substring-after(., 'US-')"/>
        </xsl:copy>
    </xsl:template>
    <xsl:template match="marc:datafield[@tag eq '040']/marc:subfield[@code = 'b'][not(. = $default-language-of-cataloging)]">
        <xsl:copy>
            <xsl:apply-templates select="@*"/>
            <xsl:value-of select="$default-language-of-cataloging"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="marc:datafield[@tag eq '041'][not(marc:subfield[@code eq 'a'][2])]"/>
    
    <!-- Yale no like the codes, but why? -->
    <xsl:template match="marc:datafield[@tag eq '041']/marc:subfield[@code eq '2']"/>
    
    <xsl:template match="marc:datafield[@tag = ('044', '049', '099')]"/>
    
    <xsl:template match="marc:datafield[@tag eq '300']/marc:subfield[@code eq 'f']">
        <xsl:copy>
            <xsl:apply-templates select="@*"/>
            <!-- do a better job for replacing these values, but this should work for now -->
            <xsl:value-of select="translate(., 'Linear Feet', 'linear feet') => replace('\(\(', '(') => replace('\)\)', ')')"/>
        </xsl:copy>
    </xsl:template>
    
    <!-- still have to address 100 - 7xx -->
    
    <xsl:template match="marc:datafield[@tag eq '852']">
        <xsl:copy>
            <xsl:apply-templates select="@*"/>
            <marc:subfield code="a">
                <xsl:value-of select="marc:subfield[@code eq 'b'][1]"/>
            </marc:subfield>
            <marc:subfield code="b">
                <xsl:value-of select="marc:subfield[@code eq 'a'][1]"/>
            </marc:subfield>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="marc:datafield[@tag eq '856']/marc:subfield[@code eq 'z']">
        <xsl:copy>
            <xsl:attribute name="code" select="'3'"/>
            <xsl:text>View a description and listing of collection contents in the finding aid</xsl:text>  
        </xsl:copy>
    </xsl:template>
    
    <!-- refactor after the full requirements are gathered -->
    <xsl:template match="marc:datafield[@tag eq '506'][marc:subfield[@code eq 'a'][matches(., '\n\n')]]">
        <xsl:copy>
            <xsl:apply-templates select="@* except @ind1"/>
            <xsl:attribute name="ind1" select="if (tokenize(marc:subfield[@code eq 'a'], '\n\n')[1] => lower-case() => contains('restricted')) then '1' else '0'"/>
            <xsl:for-each select="tokenize(marc:subfield[@code eq 'a'], '\n\n')[1]">
                <xsl:element name="subfield" namespace="http://www.loc.gov/MARC21/slim">
                    <xsl:attribute name="code" select="'a'"/>
                    <xsl:value-of select="."/>
                </xsl:element>
            </xsl:for-each>
        </xsl:copy>
        <xsl:for-each select="tokenize(marc:subfield[@code eq 'a'], '\n\n')[position() gt 1]">
            <xsl:call-template name="multi-paragraph">
                <xsl:with-param name="node" select="."/>
            </xsl:call-template>
        </xsl:for-each>
    </xsl:template>
    
    <xsl:template name="multi-paragraph">
        <xsl:param name="node"/>
        <xsl:element name="datafield" namespace="http://www.loc.gov/MARC21/slim">
            <xsl:attribute name="tag" select="'506'"/>
            <xsl:attribute name="ind1" select="if (contains(lower-case($node), 'restricted')) then '1' else '0'"/>
            <xsl:attribute name="ind2" select="' '"/>
            <xsl:element name="subfield" namespace="http://www.loc.gov/MARC21/slim">
                <xsl:attribute name="code" select="'a'"/>
                <xsl:value-of select="$node"/>
            </xsl:element>
        </xsl:element>
    </xsl:template>
    
    <xsl:template match="marc:subfield[@code eq '2'][starts-with(., 'local_')]">
        <xsl:copy>
            <xsl:apply-templates select="@*"/>
            <xsl:text>local</xsl:text>
        </xsl:copy>
    </xsl:template>
    
</xsl:stylesheet>
