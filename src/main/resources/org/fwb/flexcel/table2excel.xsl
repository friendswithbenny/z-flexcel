<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
		xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
	
	<xsl:param name="dateStyleIndex" />
	<xsl:param name="date1904" />
	
	<xsl:variable name="mils1900" select="-2209057200000" />
	<xsl:variable name="mils1904" select="-2082826800000" />
	
	<xsl:variable name="epoch">
		<xsl:choose>
			<xsl:when test="$date1904"><xsl:value-of select="$mils1904" /></xsl:when>
			<xsl:otherwise><xsl:value-of select="$mils1900" /></xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	
	<xsl:template name="convertDate">
		<xsl:param name="d" />
		<xsl:if test="$d">
			<xsl:value-of select="($d - $epoch) div (1000 * 60 * 60 * 24)" />
		</xsl:if>
	</xsl:template>
	
	<xsl:template match="table">
		<worksheet>
			<sheetViews>
				<sheetView tabSelected="0" workbookViewId="0">
					<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozenSplit" />
					<selection pane="bottomLeft" activeCell="A2" sqref="A2" />
				</sheetView>
			</sheetViews>
			<sheetFormatPr defaultRowHeight="15" />
			<sheetData>
				<xsl:apply-templates select="thead/tr|tbody/tr" />
			</sheetData>
		</worksheet>
	</xsl:template>
	
	<xsl:template match="tr">
		<row>
			<xsl:apply-templates select="td|th" />
		</row>
	</xsl:template>
	
	<xsl:template match="td|th">
		<xsl:variable name="col" select="position()" />
		<xsl:variable name="type" select="translate(/table/thead/tr/th[position()=$col]/@type, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
		<c>
			<xsl:choose>
				<xsl:when test="self::th or contains($type, 'char') or $type='text'">
					<xsl:attribute name="t">inlineStr</xsl:attribute>
					<is>
						<t>
							<xsl:apply-templates />
						</t>
					</is>
				</xsl:when>
				<xsl:when test="starts-with($type, 'bool')">
					<xsl:attribute name="t">b</xsl:attribute>
					<v>
						<xsl:value-of select="number(text() = 'true')" />
					</v>
				</xsl:when>
				<xsl:when test="$type='date'">
					<xsl:attribute name="s">
						<xsl:value-of select="$dateStyleIndex" />
					</xsl:attribute>
					<v>
						<xsl:call-template name="convertDate">
							<xsl:with-param name="d" select="node()" />
						</xsl:call-template>
					</v>
				</xsl:when>
				<xsl:when test="starts-with($type, 'int') or starts-with($type, 'num')">
					<v>
						<xsl:apply-templates />
					</v>
				</xsl:when>
				<xsl:when test="$type='clob'">
					<xsl:attribute name="t">inlineStr</xsl:attribute>
					<is>
						<t>
							<xsl:apply-templates />
						</t>
					</is>
				</xsl:when>
				<!-- default to string -->
				<!--xsl:otherwise>
					<xsl:attribute name="t">inlineStr</xsl:attribute>
					<is>
						<t>
							<xsl:apply-templates />
						</t>
					</is>
				</xsl:otherwise-->
				<!-- for debugging -->
				<xsl:otherwise>
					<xsl:attribute name="t">inlineStr</xsl:attribute>
					<is>
						<t>
							<xsl:apply-templates />
							<xsl:text> (</xsl:text>
							<xsl:value-of select="$type" />
							<xsl:text>)</xsl:text>
						</t>
					</is>
				</xsl:otherwise>
				<!-- used to default to numeric -->
				<!--xsl:otherwise>
					<v>
						<xsl:apply-templates />
					</v>
				</xsl:otherwise-->
			</xsl:choose>
		</c>
	</xsl:template>
</xsl:stylesheet>