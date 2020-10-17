<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:map="http://www.w3.org/2005/xpath-functions/map"
    xmlns:math="http://www.w3.org/2005/xpath-functions/math"
    xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl" xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:html="http://www.w3.org/TR/REC-html40" xmlns:marc="http://www.loc.gov/MARC21/slim"
    xmlns:mdc="http://mdc" exclude-result-prefixes="#all" version="3.0">
    <xsl:output method="xml" indent="true" encoding="UTF-8"/>
    <xsl:strip-space elements="*"/>
    
    <!-- to do:
        
        also need a way to indicate combined accessions.  
             
        Ellipsis (…) is identified by voyager as an unrecogized character, so I re-entered in voyager. Use a character replcement for this.

        review the punctuation updates.
        
        add a sample schematron file.... and add it to the pipeline.
        
        -->
    
    <xsl:param name="leader_05" select="'n'"/>
    <xsl:param name="cataloging-agency" select="'CtY-BR'"/>
    <xsl:param name="transcribing-agency" select="$cataloging-agency"/>
    <xsl:param name="language-of-cataloging" select="'eng'"/>
    <xsl:param name="location_852_a" select="'Beinecke Rare Book and Manuscript Library, Yale University, New Haven, CT'"/>
    <xsl:param name="default_access_note" select="'This material is open for research.'"/>
    <xsl:param name="repo_for_automation_note" select="'Beinecke Manuscript Unit'"/>
    <xsl:param name="conversion_date" select="current-date()"/>
    <xsl:param name="automation_note" select="$repo_for_automation_note || ' bulk cataloging workflow, ' || format-date($conversion_date, '[Y] [MNn,*-3]', 'en', (), ()) || '.'"/>
    <!-- not used currently, but could be utilized if we decide to combine the date one and date two fields -->
    <xsl:param name="date-separator" select="' - '"/>
    <xsl:param name="default_008_place_of_publication" select="'ctu'"/>

    <xsl:import href="Mapping-file.xsl"/>

    <!-- change this to a named templates? -->
    <xsl:function name="mdc:create-marc-datafield" as="node()">
        <xsl:param name="tag"/>
        <xsl:param name="ind1"/>
        <xsl:param name="ind2"/>
        <xsl:param name="first_subfield_code"/> <!-- might be blank, when we're passing a blob that needs to be tokenized? -->
        <xsl:param name="first_subfield_text"/>
        <marc:datafield tag="{$tag}" ind1="{$ind1}" ind2="{$ind2}">
            <xsl:sequence select="mdc:create-marc-subfield($first_subfield_code, $first_subfield_text)"/>
        </marc:datafield>
    </xsl:function>
    
    <!-- could combine these next two functions, but not sure the best way to do that just yet -->
    <xsl:function name="mdc:create-marc-subfield" as="node()">
        <xsl:param name="code"/>
        <xsl:param name="text"/>
        <marc:subfield code="{normalize-space($code)}">
            <xsl:value-of select="normalize-space($text)"/>
        </marc:subfield>
    </xsl:function>
        
    <!-- creator, title, dates of creation...  notes...  all could have assumed subfield a's. -->
    <xsl:function name="mdc:tokenize-subfields-with-implicit-first-subfield" as="node()*">
        <xsl:param name="first-subfield"/>
        <xsl:param name="text-to-tokenize"/>
        <xsl:variable name="subfields" select="tokenize($text-to-tokenize, '\$')"/>
        <xsl:for-each select="$subfields">
            <xsl:choose>
                <xsl:when test="position() = 1">
                    <marc:subfield code="{$first-subfield}">
                        <xsl:value-of select="normalize-space(.)"/>
                    </marc:subfield>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:variable name="code" select="substring(., 1, 1)"/>
                    <marc:subfield code="{$code}">
                        <xsl:value-of select="normalize-space(substring(., 2))"/>
                    </marc:subfield>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:for-each>
    </xsl:function>
    
    <xsl:function name="mdc:tokenize-subfields" as="node()*">
        <xsl:param name="input"/>
        <xsl:variable name="subfield" select="tokenize($input, ' \$')"/>
        <xsl:for-each select="$subfield[position() > 1]">
            <xsl:variable name="code" select="substring(., 1, 1)"/>
            <marc:subfield code="{$code}">
                <xsl:value-of select="normalize-space(substring(., 2))"/>
            </marc:subfield>
        </xsl:for-each>
    </xsl:function>

    <xsl:template match="ss:Workbook">
        <marc:collection xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/MARC21/slim http://www.loc.gov/standards/marcxml/schema/MARC21slim.xsd">
            <xsl:apply-templates select="ss:Worksheet[starts-with(@ss:Name, 'Cataloging')][1]/ss:Table"/>
        </marc:collection>
    </xsl:template>

    <xsl:template match="ss:Table">
        <!-- ignore the first two header rows -->
        <xsl:apply-templates select="ss:Row[position() gt 2][ss:Cell/ss:Data]"/>
    </xsl:template>

    <xsl:template match="ss:Row">
        <!-- dynamically create variables based on config file or dervied from incoming NamedCells???  well, yes, we should... but for now... -->
        <xsl:variable name="accession_number" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Accession_number']"/>
        <xsl:variable name="call_number_prefix" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Call_number_sequence']"/>
        <xsl:variable name="call_number_suffix" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Unique']"/>
        <xsl:variable name="bib_level" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Bibliographic_level']"/>
        <xsl:variable name="rules" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Cataloging_Source_040_e']"/>
        <xsl:variable name="_1xx" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = '_1xx']"/>
        <xsl:variable name="Creator_1xx" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Creator_1xx']"/>
        <xsl:variable name="_1xx_e" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = '_1xx_e']"/>
        <xsl:variable name="_1xx_0" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = '_1xx_0']"/>
        <xsl:variable name="title" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Title_245']"/> 
        <xsl:variable name="material_type" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Material_type_245']"/>
        <xsl:variable name="title_statement" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Statement_245']"/>
        <xsl:variable name="Date_of_Creation" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Date_of_Creation']"/>
        <xsl:variable name="Date_one_008" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Date_one_008']"/>
        <xsl:variable name="Date_two_008" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Date_two_008']"/>
        <xsl:variable name="Place_of_Creation_264" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Place_of_Creation_264']"/>
        <xsl:variable name="Place_code_008" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Place_code_008']"/>
        <xsl:variable name="Extent_300" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Extent_300']"/>
        <xsl:variable name="Additional_Extent_300" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Additional_Extent_300']"/>
        <xsl:variable name="Physical_Details_300" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Physical_Details_300']"/>
        <xsl:variable name="Extent_Dimensions_300" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Extent_Dimensions_300']"/>
        <xsl:variable name="Content_336" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Content_Type_336']"/>
        <xsl:variable name="Carrier_338" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Carrier_Type_338']"/>
        <xsl:variable name="Arrangement_351" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Arrangement_351']"/>
        <xsl:variable name="Access_506" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Access_Restrictions_506']"/>
        <xsl:variable name="Bioghist_545" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Biographical_Note_545']"/>
        <xsl:variable name="Scope_520" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Scope_and_Contents_520']"/>
        <xsl:variable name="General_500" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'General_Note_500']"/>
        <xsl:variable name="Title_500" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Title_Source_Note_500']"/>
        <xsl:variable name="Language_546" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Language_of_Materials_546']"/>         
        <xsl:variable name="Language_code_008" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Language_code_008']"/>
        <xsl:variable name="provenance_note" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Provenance_Note_561']"/>
        <xsl:variable name="accession_type" select="(ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Accession_Type_561'], 0)[1]"/>
        <xsl:variable name="source" select="(ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Source_561'], 0)[1]"/>
        <xsl:variable name="fund" select="(ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Fund_561'], 0)[1]"/>
        <xsl:variable name="acquistion_years" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Years_of_Acquisition_561']"/>
        <xsl:variable name="citation" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Preferred_Citation_524']"/>
        <xsl:variable name="Geographic_651" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Geographic_651']"/>
        <xsl:variable name="Geographic_subfields_651" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Geographic_subfields_651']"/>
        <xsl:variable name="Genre_Term_655" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Genre_Term_655']"/>
        <xsl:variable name="Additional_MARC" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Additional_MARC_Fields']"/>
        <xsl:variable name="location_code_1" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Location_code']"/>
        <xsl:variable name="location_code_2" select="ss:Cell[ss:Data[normalize-space()]][ss:NamedCell/@ss:Name = 'Additional_Location_code']"/>
        <!-- not checking out the Public_URL column just yet, but will need that for ASpace updates....  once we can get bibids back from Voyager -->
        
        <!--great candidate for try / catch here.  but for now, let's just switch to https due to change on 2020-08-06 at LoC  --> 
        <xsl:variable name="marcxml" select=" if ($_1xx_0) then document(replace(normalize-space($_1xx_0),  'http://', 'https://') || '.marcxml.xml') else 0"/> 
        
        <xsl:variable name="born-digital-006" select="some $text in ($Genre_Term_655, $Additional_MARC) satisfies contains(lower-case($text), 'born digital')"/>
        <xsl:variable name="drawings-006" select="some $text in ($Genre_Term_655, $Additional_MARC) satisfies contains(lower-case($text), 'drawings')"/>
        <xsl:variable name="manuscript-maps-006" select="some $text in ($Genre_Term_655, $Additional_MARC) satisfies contains(lower-case($text), 'manuscript maps')"/>
        <xsl:variable name="photographs-006" select="some $text in ($Genre_Term_655, $Additional_MARC) satisfies contains(lower-case($text), 'photographs')"/>
        <marc:record>
            <marc:leader>
                <!-- MARC Edit will  add the record length, so we keep that set to 00000 for the time being .-->
                <!-- add the leader 06 field to a mapping file? -->
                <xsl:variable name="leader_06" select="if ($rules = 'dcrmg') then 'k' else if ($rules = 'dcrmc') then 'f' else if ($rules = 'dcrmm') then 'd' else 'p'"/>
                <!-- change this once we figure out how folks should indicate this field -->
                <xsl:variable name="leader_07" select="if ($rules =('dacs', 'dcrmmss')) then 'c' else if ($bib_level) then substring(normalize-space($bib_level), 1, 1) else 'c'"/>
                <xsl:variable name="leader_08" select="if ($rules = ('dcrmc', 'dcrmg')) then ' ' else 'a'"/>
                <xsl:variable name="leader_end" select="' 2200277 i 4500'"/>
                <xsl:value-of select="'00000' || $leader_05 || $leader_06 || $leader_07 || $leader_08 || $leader_end" />
            </marc:leader>
            <xsl:call-template name="control-fields">
                <xsl:with-param name="rules" select="$rules"/>
                <xsl:with-param name="born-digital-006" select="$born-digital-006"/>
                <xsl:with-param name="drawings-006" select="$drawings-006"/>
                <xsl:with-param name="manuscript-maps-006" select="$manuscript-maps-006"/>
                <xsl:with-param name="photographs-006" select="$photographs-006"/>       
                <xsl:with-param name="Date_one_008" select="$Date_one_008"/>
                <xsl:with-param name="Date_two_008" select="$Date_two_008"/>
                <xsl:with-param name="Place_code_008" select="$Place_code_008"/>
                <xsl:with-param name="Language_code_008" select="$Language_code_008"/>
            </xsl:call-template>
            <xsl:call-template name="cataloging_040">
                <xsl:with-param name="rules" select="$rules"/>
            </xsl:call-template>
            <xsl:if test="$Creator_1xx or $_1xx_0">
                <xsl:call-template name="creator_1xx">
                    <xsl:with-param name="_1xx" select="$_1xx"/>
                    <xsl:with-param name="Creator_1xx" select="$Creator_1xx"/>
                    <xsl:with-param name="_1xx_e" select="$_1xx_e"/>
                    <xsl:with-param name="_1xx_0" select="$_1xx_0"/>
                    <xsl:with-param name="marcxml" select="$marcxml"/>
                    <xsl:with-param name="rules" select="$rules"/>
                </xsl:call-template>
            </xsl:if>
            <xsl:call-template name="title_245">
                <xsl:with-param name="title" select="$title"/>
                <xsl:with-param name="ind1" select="if ($Creator_1xx) then '1' else '0'"/>
                <xsl:with-param name="dacs_date" select="if ($rules eq 'dacs') then $Date_of_Creation else 0"/>
                <xsl:with-param name="material_type" select="$material_type"/>
                <xsl:with-param name="title_statement" select="$title_statement"/>
            </xsl:call-template>
            <xsl:if test="$rules = ('dcrmmss', 'dcrmg', 'dcrmc')">
                <xsl:call-template name="publication_264">
                    <xsl:with-param name="place" select="$Place_of_Creation_264"/>
                    <xsl:with-param name="date" select="$Date_of_Creation"/>
                </xsl:call-template>
            </xsl:if>
            <xsl:if test="$Extent_300">
                <xsl:call-template name="extent_300">
                    <xsl:with-param name="subfield_300_a" select="$Extent_300"/>
                    <xsl:with-param name="subfield_300_f" select="$Additional_Extent_300"/>
                    <xsl:with-param name="subfield_300_b" select="$Physical_Details_300"/>
                    <xsl:with-param name="subfield_300_c" select="$Extent_Dimensions_300"/>
                </xsl:call-template>               
            </xsl:if>
            <xsl:if test="$citation">
                <xsl:call-template name="citation_524">
                    <xsl:with-param name="title" select="$title"/>
                    <xsl:with-param name="citation" select="$citation"/>
                    <xsl:with-param name="rules" select="$rules"/>
                </xsl:call-template>
            </xsl:if>
            <xsl:call-template name="ownership_561">
                <xsl:with-param name="provenance_note" select="$provenance_note"/>
                <!-- the following 4 params are required, right?  if so, how should we enforce that?  
                    should we assume a default value is missing, fail the transformation, etc. -->
                <xsl:with-param name="accession_type" select="$accession_type"/>
                <xsl:with-param name="source" select="$source"/>
                <xsl:with-param name="acquisition_years" select="$acquistion_years"/>
                <xsl:with-param name="fund" select="$fund"/>
            </xsl:call-template>

            <xsl:if test="$Geographic_651">
                <xsl:sequence select="mdc:create-marc-datafield('651', ' ', '0', 'a', $Geographic_651 || ' ' || $Geographic_subfields_651)"/>
            </xsl:if>
            
            <xsl:sequence select="mdc:create-marc-datafield('852', ' ', ' ', 'a', $location_852_a)"/>
          
            <!-- all of this is a bit silly right now, but once we get all of the requirements right we can go back and refactor.-->
            <xsl:if test="$location_code_1">
                <xsl:call-template name="holdings_951">
                    <xsl:with-param name="location_code" select="$location_code_1"/>
                    <xsl:with-param name="call_number_prefix" select="$call_number_prefix"/>
                    <xsl:with-param name="call_number_suffix" select="$call_number_suffix"/>
                    <xsl:with-param name="automation_note" select="$automation_note"/>
                </xsl:call-template>
            </xsl:if>
            
            <xsl:if test="$location_code_2">
                <xsl:call-template name="holdings_951">
                    <xsl:with-param name="location_code" select="$location_code_2"/>
                    <xsl:with-param name="call_number_prefix" select="$call_number_prefix"/>
                    <xsl:with-param name="call_number_suffix" select="$call_number_suffix"/>
                    <xsl:with-param name="automation_note" select="$automation_note"/>
                </xsl:call-template>
            </xsl:if>
            
            <!-- need to update this process in case the 952 needs to record more than one accession ID.   -->
            <xsl:sequence select="mdc:create-marc-datafield('952', ' ', ' ', 'a', $accession_number)"/>

            <xsl:if test="$Content_336 and not($rules='dacs')">
                <marc:datafield tag="336" ind1=" " ind2=" ">
                    <xsl:sequence select="mdc:create-marc-subfield('a', $Content_336)"/>
                    <xsl:sequence select="mdc:create-marc-subfield('2', 'rdacontent')"/>
                </marc:datafield>
            </xsl:if>
            <xsl:if test="$Carrier_338 and not($rules='dacs')">
                <marc:datafield tag="337" ind1=" " ind2=" ">
                    <xsl:sequence select="mdc:create-marc-subfield('a', 'unmediated')"/>
                    <xsl:sequence select="mdc:create-marc-subfield('2', 'rdamedia')"/>
                </marc:datafield>
                <marc:datafield tag="338" ind1=" " ind2=" ">
                    <xsl:sequence select="mdc:create-marc-subfield('a', $Carrier_338)"/>
                    <xsl:sequence select="mdc:create-marc-subfield('2', 'rdacarrier')"/>
                </marc:datafield>
            </xsl:if>
            <xsl:if test="$Arrangement_351">
                <xsl:variable name="subfield" select="if (matches($Arrangement_351, 'organized|series', 'i')) then 'a' else 'b'"/>
                <xsl:sequence select="mdc:create-marc-datafield('351', ' ', ' ', $subfield, $Arrangement_351)"/>
            </xsl:if>
            <xsl:sequence select="mdc:create-marc-datafield('506', '0', ' ', 'a', $default_access_note)"/>
            <!-- adjust as requested by feedback based on how folks formulate these notes -->
             <xsl:variable name="Access_506_subfield" select="if (matches($Access_506, 'container|box|folder|roll|broadside')) then '3' else 'a'"/>
            <!-- do we need to change ind1 based on the text? -->
            <xsl:if test="$Access_506">
                 <xsl:sequence select="mdc:create-marc-datafield('506', '1', ' ', $Access_506_subfield, $Access_506)"/>
             </xsl:if>
            <xsl:if test="$Bioghist_545">
                <xsl:sequence select="mdc:create-marc-datafield('545', ' ', ' ', 'a', $Bioghist_545)"/>
            </xsl:if>
            <xsl:if test="$Scope_520">
                <xsl:sequence select="mdc:create-marc-datafield('520', ' ', ' ', 'a', $Scope_520)"/>
            </xsl:if>
            <xsl:if test="$General_500">
                <xsl:sequence select="mdc:create-marc-datafield('500', ' ', ' ', 'a', $General_500)"/>
            </xsl:if>
            <xsl:if test="$Title_500">
                <xsl:sequence select="mdc:create-marc-datafield('500', ' ', ' ', 'a', $Title_500)"/>
            </xsl:if>
            <xsl:if test="$Language_546">
                <xsl:sequence select="mdc:create-marc-datafield('546', ' ', ' ', 'a', $Language_546)"/>
            </xsl:if>
            <xsl:if test="$Genre_Term_655">
                <marc:datafield tag="655" ind1=" " ind2="7">
                    <xsl:choose>
                        <xsl:when test="contains($Genre_Term_655, '$')">
                            <xsl:sequence select="mdc:tokenize-subfields-with-implicit-first-subfield('a', $Genre_Term_655)"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:sequence select="mdc:create-marc-subfield('a', $Genre_Term_655)"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </marc:datafield>
            </xsl:if>
            <!-- adding this loop in case folks want to add mutiple columns.  if so, then they'll need to make sure that the column name extends to each column, just like the header -->
            <xsl:for-each select="$Additional_MARC">
                <xsl:variable name="marc_field" select="tokenize(.,  '&#10;')"/>
                <!-- if any tokenized string doesn't have a tag, or is just an empty line, skip it. -->
                <xsl:for-each select="$marc_field[string-length(.) > 3]">
                    <xsl:variable name="tag" select="substring(normalize-space(.), 1, 3)"/>
                    <xsl:variable name="ind1" select="if (substring(normalize-space(.), 5, 1) eq '_') then ' ' else substring(normalize-space(.), 5, 1)"/>
                    <xsl:variable name="ind2" select="if (substring(normalize-space(.), 6, 1) eq '_') then ' ' else substring(normalize-space(.), 6, 1)"/>
                    <marc:datafield tag="{$tag}" ind1="{$ind1}" ind2="{$ind2}">
                        <xsl:variable name="subfield" select="tokenize(., ' \$')"/>
                        <xsl:for-each select="$subfield[position() > 1]">
                            <xsl:variable name="code" select="substring(., 1, 1)"/>
                            <xsl:sequence select="mdc:create-marc-subfield($code, normalize-space(substring(., 2)))"/>
                        </xsl:for-each>
                    </marc:datafield>
                </xsl:for-each>
            </xsl:for-each>                
        </marc:record>
    </xsl:template>

    <xsl:template name="control-fields">
        <xsl:param name="rules"/>
        <xsl:param name="born-digital-006"/>
        <xsl:param name="drawings-006"/>
        <xsl:param name="photographs-006"/>
        <xsl:param name="manuscript-maps-006"/>
        <xsl:param name="Date_one_008"/>
        <xsl:param name="Date_two_008"/>
        <xsl:param name="Place_code_008"/>
        <xsl:param name="Language_code_008"/>
        
        <!--
        <xsl:variable name="field006_00-04" select="if ($Genre_Term_655 eq 'Photographs') then 'kn   '
            else if ($Genre_Term_655 eq 'Drawings') then 'an   '
            else if ($Genre_Term_655 eq 'Manuscript maps') then 'a0e  '
            else if ($rules eq 'dcrmmss') then 't     ' else '     '"/>
        
        <xsl:variable name="field006_05-11" select="'       '"/>
        <xsl:variable name="field006_12-17" select="'000 0 '"/>
        
         <xsl:choose>
                <xsl:when test="$rules = 'dcrmmss'">
                    <xsl:value-of select="$field006_00-04 || $field006_05-11 || $field006_12-17"/>    
                </xsl:when>
                <xsl:when test="$rules = ('dacs')">
                    <xsl:value-of select="$field006_00-04"/>    
                </xsl:when>
                <xsl:otherwise/>
            </xsl:choose>
        -->
        
        <xsl:variable name="field008_00-05"  select="substring(string(current-date()), 3, 2) || substring(string(current-date()), 6, 2) || substring(string(current-date()), 9, 2)"/>
        <xsl:variable name="field008_06"
            select="
                if ($Date_two_008) then 'i'
                else if ($Date_one_008) then 's'
                else '|'"/>
        <!-- would it ever be possible for the Date_one or Date_two variable to be less than 4 characters?
            if so, we'll either need to add spaces, or raise an error -->
        <xsl:variable name="field008_07-10"
            select="if ($Date_one_008) then $Date_one_008 else '    '"/>
        <xsl:variable name="field008_11-14"
            select="
                if ($Date_two_008) then
                    $Date_two_008
                else
                    '    '"/>
        <xsl:variable name="field008_15-17"
            select="
                if ($rules eq 'dacs') then $default_008_place_of_publication
                else if (string-length($Place_code_008) eq 3) then
                    $Place_code_008
                    else if (string-length($Place_code_008) eq 2) then 
                    $Place_code_008 || '#'
                else
                    'xx '"/>
       <!-- dcrmm differences.  -->
        <xsl:variable name="field008_18-34" select="
            if ($rules = 'dcrmm') 
            then 'uuu _ ______n_   '
            else
            '     _           '"/>
        <xsl:variable name="field008_35-37"
            select="
                if ($Language_code_008) then
                substring($Language_code_008, 1, 3)
                else
                    '   '"/>
        <xsl:variable name="field008_38-39" select="' d'"/>

       
        <!-- 006 -->
        <!-- could have multiple 006 fields.  we'll call them here, as necessary -->
        <xsl:if test="$born-digital-006">
            <xsl:call-template name="create_006">
                <xsl:with-param name="values" select="'m    __  | |      '"/>
            </xsl:call-template>
        </xsl:if>
        <xsl:if test="$drawings-006">
            <xsl:call-template name="create_006">
                <xsl:with-param name="values" select="'knnn _     __   an'"/>
            </xsl:call-template>
        </xsl:if>
        <xsl:if test="$photographs-006">
            <xsl:call-template name="create_006">
                <xsl:with-param name="values" select="'knnn _     __   kn'"/>
            </xsl:call-template>
        </xsl:if>
        <xsl:if test="$manuscript-maps-006">
            <xsl:call-template name="create_006">
                <xsl:with-param name="values" select="'f______ a  __ 0 e_'"/>
            </xsl:call-template>
        </xsl:if>
        <xsl:if test="$rules = ('dcrmmss')">
            <xsl:call-template name="create_006">
                <xsl:with-param name="values" select="'t___________000 0_'"/>
            </xsl:call-template>
        </xsl:if>
    
        <!-- 007 -->
        <!-- will likely change how we handle the 007 field over time -->
        <xsl:if test="$rules = ('dcrmc')">
            <marc:controlfield tag="007">
                <xsl:value-of select="'mj ||nzn'"/>
            </marc:controlfield>
        </xsl:if>
        <xsl:if test="$rules = ('dcrmg')">
            <marc:controlfield tag="007">
                <xsl:value-of select="'k|||||'"/>
            </marc:controlfield>
        </xsl:if>
        <!-- 008 -->
        <marc:controlfield tag="008">
            <xsl:value-of select="$field008_00-05 || $field008_06 || $field008_07-10 || $field008_11-14 || $field008_15-17 || $field008_18-34 || $field008_35-37 || $field008_38-39"/>
        </marc:controlfield>
    </xsl:template>
    
    <xsl:template name="create_006">
        <xsl:param name="values"/>
        <marc:controlfield tag="006">
            <xsl:value-of select="$values"/>
        </marc:controlfield>
    </xsl:template>

    <xsl:template name="cataloging_040">
        <xsl:param name="rules"/>
        <marc:datafield tag="040" ind1=" " ind2=" ">
            <xsl:sequence select="mdc:create-marc-subfield('a', $cataloging-agency)"/>
            <xsl:sequence select="mdc:create-marc-subfield('b', $language-of-cataloging)"/>
            <xsl:sequence select="mdc:create-marc-subfield('c', $transcribing-agency)"/>
            <xsl:sequence select="mdc:create-marc-subfield('e', $rules)"/>
            <!-- For DCRMG and DCRMC records, the 040 should also include $e rda at the end.-->
            <xsl:if test="$rules = ('dcrmg', 'dcrmc')">
                <xsl:sequence select="mdc:create-marc-subfield('e', 'rda')"/>
            </xsl:if>
        </marc:datafield>
    </xsl:template>

    <xsl:template name="creator_1xx">
        <xsl:param name="marcxml"/>
        <xsl:param name="rules"/>
        <xsl:param name="_1xx"/>
        <xsl:param name="Creator_1xx"/>
        <xsl:param name="_1xx_e"/>
        <xsl:param name="_1xx_0"/>
        <xsl:variable name="tag" select="if ($marcxml) then $marcxml//marc:datafield[starts-with(@tag, '1')]/@tag  else substring-before($_1xx, ' ')"/>
        <xsl:variable name="ind1" select="substring-after($_1xx, ' ') => substring(1, 1)"/>
        <xsl:variable name="ind2" select="substring-after($_1xx, ' ') => substring(2, 1)"/>
        <xsl:choose>
            <!-- a bit hacky to change the marcxml prefix to marc.  probably shouldn't care, but doing that just in case there might be a downstream effect -->
            <xsl:when test="$marcxml">
                <xsl:apply-templates select="$marcxml//marc:datafield[@tag = $tag]" mode="copy">
                    <xsl:with-param name="_1xx_e" select="$_1xx_e"/>
                    <xsl:with-param name="_1xx_0" select="$_1xx_0"/>
                </xsl:apply-templates>
            </xsl:when>
            <xsl:otherwise>
                <marc:datafield tag="{$tag}">
                    <xsl:attribute name="ind1" select="if ($ind1 = ('#', '_')) then ' ' else $ind1"/>
                    <xsl:attribute name="ind2" select="if ($ind2 = ('#', '_')) then ' ' else $ind2"/>
                    <!-- turn this stuff in to a function later on.  will also need for dates... and then for the catch-all column, we're gonna have yet another approach to that. -->
                    <xsl:choose>
                        <xsl:when test="contains($Creator_1xx, '$')">
                            <!-- in this special case, as with dates, the first subfield is assumed to be subfield 'a'.  so, it's not explictly entered in the spreadsheet.
                                do we need to add a sanity check here??? -->
                            <!-- if we combine these 2 functions, we can remove this choose section -->
                            <xsl:sequence select="mdc:tokenize-subfields-with-implicit-first-subfield('a', $Creator_1xx)"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:sequence select="mdc:create-marc-subfield('a', $Creator_1xx)"/>
                        </xsl:otherwise>
                    </xsl:choose>
                    <xsl:if test="$_1xx_e and $rules = ('dcrmg', 'dcrmc')">
                        <xsl:sequence select="mdc:create-marc-subfield('e', $_1xx_e)"/>
                    </xsl:if>
                    <xsl:if test="$_1xx_0">
                        <xsl:sequence select="mdc:create-marc-subfield('0', $_1xx_0)"/>
                    </xsl:if>
                </marc:datafield>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="title_245">
        <xsl:param name="ind1"/>
        <xsl:param name="title"/>
        <xsl:param name="dacs_date"/>
        <xsl:param name="material_type"/>
        <xsl:param name="title_statement"/>
        <marc:datafield tag="245">
            <xsl:attribute name="ind1" select="$ind1"/>
            <!-- how to get the number of non-filing characters??? use a map/list? -->
            <xsl:attribute name="ind2" select="0"/>
            <xsl:choose>
                <xsl:when test="contains($title, '$')">
                    <xsl:sequence select="mdc:tokenize-subfields-with-implicit-first-subfield('a', $title)"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:sequence select="mdc:create-marc-subfield('a', $title)"/>
                </xsl:otherwise>
            </xsl:choose>
            <!-- what if someone puts a $f in the title field???...  also, how do we handle the re-ordering here??? -->
            <xsl:if test="$dacs_date">
                <xsl:choose>
                    <xsl:when test="contains($dacs_date, '$')">
                        <xsl:sequence select="mdc:tokenize-subfields-with-implicit-first-subfield('f', $dacs_date)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:sequence select="mdc:create-marc-subfield('f', $dacs_date)"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:if>
            <xsl:if test="$material_type">
                <!-- sample:
                    "printout"
                    (in the Deal-with-ISBD-issues section will make sure to have the 'space colon space' separator)
                    will also need to either add a terminal period, if the last subfield, or something else if not.
                -->
                <xsl:sequence select="mdc:create-marc-subfield('k', normalize-space($material_type))"/>
            </xsl:if>
            <xsl:if test="$title_statement">
                <xsl:sequence select="mdc:create-marc-subfield('c', normalize-space($title_statement))"/>
            </xsl:if>
        </marc:datafield>
    </xsl:template>

    <xsl:template name="publication_264">
        <xsl:param name="place"/>
        <xsl:param name="date"/>
        <marc:datafield tag="264" ind1=" " ind2="0">
            <xsl:sequence select="mdc:create-marc-subfield('a', $place)"/>
            <xsl:choose>
                <xsl:when test="contains($date, '$')">
                    <xsl:sequence select="mdc:tokenize-subfields-with-implicit-first-subfield('c', $date)"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:sequence select="mdc:create-marc-subfield('c', $date)"/>
                </xsl:otherwise>
            </xsl:choose>
        </marc:datafield>
    </xsl:template>

    <xsl:template name="extent_300">
        <xsl:param name="subfield_300_a"/>
        <xsl:param name="subfield_300_f"/>
        <xsl:param name="subfield_300_b"/>
        <xsl:param name="subfield_300_c"/>
        <marc:datafield tag="300" ind1=" " ind2=" ">
            <xsl:sequence select="mdc:create-marc-subfield('a', $subfield_300_a)"/>
            <xsl:if test="$subfield_300_f">
                <xsl:sequence select="mdc:create-marc-subfield('f', $subfield_300_f)"/>
            </xsl:if>
            <xsl:if test="$subfield_300_b">
                <xsl:sequence select="mdc:create-marc-subfield('b', $subfield_300_b)"/>
            </xsl:if>
            <xsl:if test="$subfield_300_c">
                <xsl:sequence select="mdc:create-marc-subfield('c', $subfield_300_c)"/>
            </xsl:if>
        </marc:datafield>    
    </xsl:template>

    <xsl:template name="citation_524">
        <xsl:param name="title"/>
        <xsl:param name="citation"/>
        <xsl:param name="rules"/>
        <!-- will it have trailing punctuation or not? -->
        <xsl:variable name="title-prep"
            select="if (contains($title, ': $')) then
            normalize-space(substring-before($title, ': $')) || '.'
            else if (contains($title, '$')) then
                    normalize-space(substring-before($title, ' $')) || '.'
            else if (not(matches(normalize-space($title), '[.]$'))) then
                    normalize-space($title) || '.'
                else
                    $title"/>
        <xsl:variable name="title-to-concat" select="if ($rules = ('dcrmc', 'dcrmg', 'dcrmm')) then $title-prep => translate('[]', '') else $title-prep"/>
        <marc:datafield tag="524" ind1=" " ind2=" ">
            <xsl:sequence select="mdc:create-marc-subfield('a', normalize-space($title-to-concat) || ' ' || map:get($curatorial_mapping, lower-case($citation)))"></xsl:sequence>
        </marc:datafield>
    </xsl:template>

    <xsl:template name="ownership_561">
        <xsl:param name="provenance_note"/>
        <xsl:param name="accession_type"/> <!-- default value = 0, which means no match will be found in our mapping... but we could use plain text from the spreadsheet? -->
        <xsl:param name="acquisition_years"/>
        <xsl:param name="fund"/>  <!-- default value = 0 -->
        <xsl:param name="source"/> <!-- default value = 0 -->
        <marc:datafield tag="561" ind1=" " ind2=" ">
            <marc:subfield code="a">
                <!-- assuming that the note will end with its own punctuation.  need to revisit that assumption and confirm-->
                <xsl:if test="$provenance_note">
                    <xsl:value-of select="$provenance_note || ' '"/>
                </xsl:if>
                <!-- e.g. Purchased from donor on fund, year -->
                <!-- redo this once the logic is worked out -->
                <xsl:variable name="acquisition_type_text" select="if (map:contains($accession_type_mapping, $accession_type)) then map:get($accession_type_mapping, $accession_type) else if ($accession_type) then upper-case(substring($accession_type,1,1)) || substring($accession_type,2) || ' of' else 0"/>
                <xsl:variable name="fund_text" select="if (map:contains($fund_mapping, $fund)) then map:get($fund_mapping, $fund) else 0"/>
                <xsl:choose>
                    <xsl:when test="$acquisition_type_text and $source and $fund_text">
                        <xsl:value-of select="$acquisition_type_text || ' ' || $source
                            || ' on the '
                            || $fund_text"/>
                    </xsl:when>
                    <xsl:when test="$acquisition_type_text and $source  and not($fund_text)">
                        <xsl:value-of select="$acquisition_type_text || ' ' || $source" />
                    </xsl:when>
                    <xsl:when test="$acquisition_type_text and not($source)  and not($fund_text)">
                        <xsl:value-of select="$acquisition_type_text || ' an unknown source'" />
                    </xsl:when>
                    <xsl:when test="$acquisition_type_text and not($source)  and $fund_text">
                        <xsl:value-of select="$acquisition_type_text || 'an unknown source on the ' || $fund_text" />
                    </xsl:when>
                    <xsl:when test="not($acquisition_type_text) and not($source)  and not($fund_text)">
                        <xsl:text>Source unknown</xsl:text>
                    </xsl:when>
                    <xsl:when test="not($acquisition_type_text) and not($source)  and $fund_text">
                        <xsl:value-of select="'Purchased on the ' || $fund_text"/>
                    </xsl:when>
                    <xsl:when test="not($acquisition_type_text) and $source  and $fund_text">
                        <xsl:value-of select="'Purchased from  ' || $source || ' on the ' || $fund_text"/>
                    </xsl:when>
                    <xsl:when test="not($acquisition_type_text) and $source  and not($fund_text)">
                        <xsl:value-of select="'Acquired from  ' || $source"/>
                    </xsl:when>
                </xsl:choose>
                <!-- we'll handle trailing punctuation elsewhere, if need be -->
                <xsl:if test="$acquisition_years">
                    <xsl:value-of select="', ' || $acquisition_years"/>
                </xsl:if>
            </marc:subfield>
        </marc:datafield>
    </xsl:template>
  
    <xsl:template name="holdings_951">
        <xsl:param name="location_code"/>
        <xsl:param name="call_number_prefix"/>
        <xsl:param name="call_number_suffix"/>
        <xsl:param name="automation_note"/>
        <!-- 
        $b: Location code (nr)
        $h: Call number sequence (nr)
        $i: Call number unique ID (nr)
        $x: Nonpublic note (r)
        $z: Public note (r)
        $m: Barcode (nr)
        $n: Permloc (nr)
        $o: Item type (nr)
        $p: Enum (nr)
        $q: Statistical category (r)
        
            where nr = non-repeatable, and r = repeatable
        -->
        <marc:datafield tag="951" ind1="8" ind2="0">
            <xsl:sequence select="mdc:create-marc-subfield('b', $location_code)"/>
            <xsl:sequence select="mdc:create-marc-subfield('h', $call_number_prefix)"/>
            <xsl:sequence select="mdc:create-marc-subfield('i', $call_number_suffix)"/>
            <xsl:sequence select="mdc:create-marc-subfield('x', $automation_note)"/>
        </marc:datafield>
    </xsl:template> 
    
    <xsl:template match="@*" mode="copy">
        <xsl:copy copy-namespaces="no">
            <xsl:value-of select="."/>
        </xsl:copy>
    </xsl:template>

    <xsl:template match="*" mode="copy">
        <xsl:param name="_1xx_e"/>
        <xsl:param name="_1xx_0"/>
        <xsl:element name="marc:{local-name()}">
            <xsl:apply-templates select="node() | @*" mode="copy"/>
            <xsl:if test="..[marc:datafield] and $_1xx_e">
                <xsl:sequence select="mdc:create-marc-subfield('e', $_1xx_e)"/>
            </xsl:if>
            <xsl:if test="..[marc:datafield]">
                <xsl:sequence select="mdc:create-marc-subfield('0', $_1xx_0)"/>
            </xsl:if>
        </xsl:element>
    </xsl:template>

</xsl:stylesheet>
