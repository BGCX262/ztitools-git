<?xml version="1.0" encoding="ISO-8859-1" standalone="yes"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#" xmlns:rss="http://purl.org/rss/1.0/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:rss091="http://purl.org/rss/1.0/modules/rss091/" version="1.0" rdf:dummy-for-xmlns="" rss:dummy-for-xmlns="" dc:dummy-for-xmlns="" rss091:dummy-for-xmlns="">

<xsl:output method="html"/>

	<xsl:template match="/">
		<xsl:apply-templates select="/rdf:RDF/rss:channel" mode="margin"/>
		<xsl:apply-templates select="/rdf:RDF/rss:item" mode="margin"/>
		<xsl:apply-templates select="/rdf:RDF/rss:textinput" mode="body"/>
		<xsl:apply-templates select="/rdf:RDF/rss:image"/>
	</xsl:template>
		
		
	<xsl:template match="rss:channel" mode="margin">
		<b><a href="{rss:link}" target="_blank"><xsl:value-of select="rss:title"/></a></b>
	</xsl:template>
	<xsl:template match="rss:item" mode="margin">
		<li><a href="{rss:link}" target="_blank"><xsl:value-of select="rss:title"/></a></li>
	</xsl:template>
	
	
	<xsl:template match="rss:channel" mode="body">
		<b><a href="{rss:link}" target="_blank"><xsl:value-of select="rss:title"/></a></b><br/>
		<xsl:value-of select="rss:description"/>
		<hr/>
	</xsl:template>	
	<xsl:template match="rss:item" mode="body">
		<a href="{rss:link}" target="_blank">
			<xsl:value-of select="rss:title"/>
		</a>
		<br/>
		<xsl:value-of select="rss:description"/>
		<p/>
	</xsl:template>
	
	<xsl:template match="rss:textinput" mode="body">
		<form action="{rss:link}" method="post">
		<xsl:value-of select="rss:description"/><br/>
		<input type="text" name="{rss:query}"/><br/><input type="submit" value="{rss:title}"/>
		</form>
	</xsl:template>
	
	<xsl:template match="rss:image" >
		<div align="right">
		<a href="{rss:link}" target="_blank">
			<img src="{rss:url}" alt="{rss:title}" border="0">
				<xsl:if test="string-length(width)&gt;0"><xsl:attribute name="width"><xsl:value-of select="width"/></xsl:attribute></xsl:if>
				<xsl:if test="string-length(height)&gt;0"><xsl:attribute name="height"><xsl:value-of select="height"/></xsl:attribute></xsl:if>
			</img>
		</a><br/>
		<xsl:value-of select="rss:title"/>
		</div>
	</xsl:template>
	
</xsl:stylesheet>