<?xml version="1.0" encoding="ISO-8859-1" standalone="yes"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">

<xsl:output method="html" encoding="ISO-8859-1"/>

	<xsl:template match="/">
		<xsl:apply-templates select="//channel" mode="margin"/>
		<xsl:apply-templates select="//item" mode="margin"/>
		<xsl:apply-templates select="//textinput" mode="margin"/>
		<xsl:apply-templates select="//image"/>
	</xsl:template>
		
		
	<xsl:template match="channel" mode="margin">
		<b><a href="{link}" target="_blank"><xsl:value-of select="title" disable-output-escaping="yes"/></a></b>
	</xsl:template>
	<xsl:template match="item" mode="margin">
		<li><a href="{link}" target="_blank"><xsl:value-of select="title" disable-output-escaping="yes"/></a></li>
	</xsl:template>	
	
	<xsl:template match="textinput" mode="margin">
		<table width="100%"><tr>		
		<form action="{link}" method="post"><td align="right">
		<xsl:value-of select="description" disable-output-escaping="yes"/><br/>
		<input type="text" name="{query}" size="12"/><br/><input type="submit" value="{title}"/>
		</td></form></tr></table>
	</xsl:template>		
	
	<xsl:template match="channel" mode="body">
		<b><a href="{link}" target="_blank"><xsl:value-of select="title" disable-output-escaping="yes"/></a></b><br/>
		<xsl:value-of select="description" disable-output-escaping="yes"/>
		<hr/>
	</xsl:template>	
	<xsl:template match="item" mode="body">
		<a href="{link}" target="_blank">
			<xsl:value-of select="title"/>
		</a>
		<br/>
		<xsl:value-of select="description"/>
		<p/>
	</xsl:template>		
	<xsl:template match="textinput" mode="body">
		<table><tr>		
		<form action="{link}" method="post"><td>
		<xsl:value-of select="description" disable-output-escaping="yes"/><br/>
		<input type="text" name="{query}"/><input type="submit" value="{title}"/>
		</td></form></tr></table>
	</xsl:template>
	
	
	
	<xsl:template match="image" >
		<xsl:if test="string-length(url)&gt;0">
		<div align="right">
		<a href="{link}" target="_blank">
			<img src="{url}" alt="{title}" border="0">
				<xsl:if test="string-length(width)&gt;0"><xsl:attribute name="width"><xsl:value-of select="width"/></xsl:attribute></xsl:if>
				<xsl:if test="string-length(height)&gt;0"><xsl:attribute name="height"><xsl:value-of select="height"/></xsl:attribute></xsl:if>
			</img>						
		</a><br/>
		<xsl:value-of select="title" disable-output-escaping="yes"/>
		</div></xsl:if>
	</xsl:template>
	
</xsl:stylesheet>