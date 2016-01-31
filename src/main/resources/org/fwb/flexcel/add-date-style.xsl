<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
		xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
		xmlns:xl="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
	
	<!-- very strange, given the milliseconds ".000" that a cell's display can hold such data but the formula bar actually does not! -->
	<xsl:param name="formatCode">yyyy/mm/dd\ hhmm:ss.000</xsl:param>
	
	<!--xsl:variable name="numFmtId" select="1 + count(/xl:styleSheet[not(xl:numFmts)])*163 + count(/xl:styleSheet[xl:numFmts])*/xl:styleSheet/xl:numFmts/xl:numFmt[not(preceding-sibling::xl:numFmt/@numFmtId > @numFmtId or following-sibling::xl:numFmt/@numFmtId > @numFmtId)]/@numFmtId"-->
	<xsl:variable name="numFmtId">
		<xsl:choose>
			<xsl:when test="/xl:styleSheet/xl:numFmts">
				<xsl:value-of select="1 + /xl:styleSheet/xl:numFmts/xl:numFmt[not(preceding-sibling::xl:numFmt/@numFmtId > @numFmtId or following-sibling::xl:numFmt/@numFmtId > @numFmtId)]/@numFmtId" />
			</xsl:when>
			<xsl:otherwise>164</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	
	<!-- identity -->
	<xsl:template match="@*|node()">
		<xsl:copy>
			<xsl:apply-templates select="@*|node()" />
		</xsl:copy>
	</xsl:template>
	
	<!-- add -->
	
	<xsl:template match="xl:cellXfs">
		<xsl:copy>
			<xsl:attribute name="count">
				<xsl:value-of select="@count + 1" />
			</xsl:attribute>
			
			<xsl:apply-templates select="@*|node()" />
			
			<xf numFmtId="{$numFmtId}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" />
			<!--xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" /-->
			<!--xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/-->
			<!--xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/-->
		</xsl:copy>
	</xsl:template>
	<xsl:template match="xl:cellXfs/@count" />
	
	<xsl:template match="xl:styleSheet">
		<xsl:copy>
			<xsl:apply-templates select="@*" />
			
			<numFmts>
				<xsl:attribute name="count">
					<xsl:choose>
						<xsl:when test="xl:numFmts">
							<xsl:value-of select="1 + xl:numFmts/@count" />
						</xsl:when>
						<xsl:otherwise>1</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				
				<xsl:apply-templates select="xl:numFmts/node()" />
				
				<numFmt numFmtId="{$numFmtId}" formatCode="{$formatCode}" />
			</numFmts>
			
			<xsl:apply-templates />
		</xsl:copy>
	</xsl:template>
	<!-- filter -->
	<xsl:template match="xl:numFmts" />
	
	<xsl:template mode="numFmtId" match="numFmt[preceding-sibling::numFmt/@numFmtId > @numFmtId or following-sibling::numFmt/@numFmtId > @numFmtId]" />
	<xsl:template mode="numFmtId" match="numFmt">
		<xsl:value-of select="1 + @numFmtId" />
	</xsl:template>
</xsl:stylesheet>