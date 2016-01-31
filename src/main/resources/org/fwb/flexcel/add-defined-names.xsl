<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
		xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
		xmlns:xl="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
		xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
	
	<!-- identity -->
	<xsl:template match="@*|node()">
		<xsl:copy>
			<xsl:apply-templates select="@*|node()" />
		</xsl:copy>
	</xsl:template>
	
	<!-- add -->
	<xsl:template match="xl:workbook[not(xl:definedNames)]">
		<xsl:copy>
			<xsl:apply-templates select="@*" />
			<xsl:apply-templates select="xl:fileVersion|xl:fileSharing|xl:workbookPr|xl:workbookProtection|xl:bookViews|xl:sheets|xl:functionGroups|xl:externalReferences" />
			<definedNames />
			<xsl:apply-templates select="xl:calcPr|xl:oleSize|xl:customWorkbookViews|xl:pivotCaches|xl:smartTagPr|xl:smartTagTypes|xl:webPublishing|xl:fileRecoveryPr|xl:webPublishObjects|xl:extLst" />
		</xsl:copy>
	</xsl:template>
	
</xsl:stylesheet>