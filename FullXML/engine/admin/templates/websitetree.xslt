<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xml:space="preserve">
	<xsl:output method="html" encoding="UTF-8" />
	
	<xsl:param name="websiteID"/>
	
	<xsl:template match="/">		
		<xsl:apply-templates  select="website/menus"/>	
	</xsl:template>
		
	<xsl:template match="menus">
		<xsl:apply-templates select="menu"/>
		
		<!-- if no menu, propose a link to insert ,  -->
		<xsl:if test="count(menu)=0">
			<a href="admin.asp?websiteID={$websiteID}&amp;action=editmenu" class="smallcommand">add menu</a>
		</xsl:if>
	</xsl:template>

	<xsl:template match="menu">
		
		<table class="menu">
			<caption><a href="admin.asp?websiteID={$websiteID}&amp;action=editmenu&amp;id={@id}"><b><xsl:apply-templates select="@name"/></b></a></caption>
			<xsl:apply-templates  select="page"/>
			
			<!-- if there is no page in the menu,  propose a link to insert -->
			<xsl:if test="count(page)=0">
				<tr><td><a href="admin.asp?websiteID={$websiteID}&amp;action=editpage&amp;menuID={@id}" class="smallcommand">+</a></td></tr>
			</xsl:if>
		</table>
		
		<!-- Insert menu -->
		<a href="admin.asp?websiteID={$websiteID}&amp;action=editmenu&amp;beforemenuID={following-sibling::menu/@id}">add menu</a><p/>
	</xsl:template>

	<!-- template that display a page in the website tree -->
	<xsl:template match="page">
		<tr>
			<td>
				<a href="admin.asp?websiteID={$websiteID}&amp;action=editpage&amp;id={@id}"><xsl:apply-templates select="@name"/></a>
				&#32;[<a href="admin.asp?websiteID={$websiteID}&amp;action=editpage&amp;menuID={parent::menu/@id}&amp;beforepageID={following-sibling::page/@id}" class="smallcommand">+</a>]
			</td>
		</tr>
	</xsl:template>
</xsl:stylesheet>
