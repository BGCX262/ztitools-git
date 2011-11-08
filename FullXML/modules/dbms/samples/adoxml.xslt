<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
 xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
 xmlns:msxsl="urn:schemas-microsoft-com:xslt"
 xmlns:rs="urn:schemas-microsoft-com:rowset" 
 xmlns:s="uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882" 
 xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" 
 xmlns:z="#RowsetSchema"
 xmlns:user="http://nourl.com/mynamespace"
 version="1.0">

<msxsl:script language="VBScript" implements-prefix="user">

Function formatDate(nodelist)
 strDate = Split(nodelist.item(0).text,"T")
 
 'formatDate = CStr(FormatDateTime(CDate(strDate(0)),vbShortDate))
 formatDate = strDate(0)
End Function

</msxsl:script>
 
<xsl:template match="/">
 <TABLE cellpadding="3" cellspacing="1" border="0">
  <TR>
  <xsl:for-each select="xml/s:Schema/s:ElementType/s:AttributeType">
   <TD><B><xsl:value-of select="@name"/></B></TD>
  </xsl:for-each>
  </TR>
  <xsl:for-each select="xml/rs:data/z:row">
   <TR>
    <xsl:for-each select="attribute::*">
     <xsl:variable name="COL" select="name(.)"/>
     <xsl:variable name="DATATYPE" select="//xml/s:Schema/s:ElementType/s:AttributeType[@name = $COL]/s:datatype/@dt:type"/>
     
     <xsl:choose>
      <xsl:when test="$DATATYPE = 'string'">
       <TD TITLE="String"><xsl:value-of select="." disable-output-escaping="yes"/></TD>
      </xsl:when>
      <xsl:when test="$DATATYPE = 'int'">
       <TD TITLE="Integer"><xsl:value-of select="." disable-output-escaping="yes"/></TD>
      </xsl:when>
      <xsl:when test="$DATATYPE = 'dateTime'">
       <TD TITLE="Date/Time"><xsl:value-of select="user:formatDate(.)" disable-output-escaping="yes"/></TD>
      </xsl:when>
      <xsl:otherwise>
       <TD TITLE="Unknown or unspecified Data Type">
        <xsl:value-of select="." disable-output-escaping="yes"/>
       </TD>
      </xsl:otherwise>
     </xsl:choose>
     
    </xsl:for-each>
   </TR>
  </xsl:for-each>
 </TABLE>
</xsl:template>

</xsl:stylesheet>