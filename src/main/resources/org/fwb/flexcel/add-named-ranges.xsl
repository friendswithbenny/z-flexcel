<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
		xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
		xmlns:xp="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
		xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
	
	<!-- identity -->
	<xsl:template match="@*|node()">
		<xsl:copy>
			<xsl:apply-templates select="@*|node()" />
		</xsl:copy>
	</xsl:template>
	
	<!-- add -->
	<xsl:template match="xp:HeadingPairs/vt:vector[not(vt:variant/vt:lpstr[text() = 'Named Ranges'])]">
		<xsl:copy>
			<xsl:attribute name="size"><xsl:value-of select="@size + 2" /></xsl:attribute>
			<xsl:apply-templates select="@baseType" />
			<xsl:apply-templates />
			<vt:variant>
				<vt:lpstr>Named Ranges</vt:lpstr>
			</vt:variant>
			<vt:variant>
				<vt:i4>0</vt:i4>
			</vt:variant>
		</xsl:copy>
	</xsl:template>
	
</xsl:stylesheet>