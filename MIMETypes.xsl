<?xml version="1.0" encoding="ISO-8859-1" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:param name="Data"></xsl:param>
	<xsl:template match="IISRegisteredMimeTypes">
		<html>
			<head>
				<style>

					Table.Report
					{
					border: 1px solid black;
					border-collapse: collapse;
					font-family: calibri;
					font-size: 14px;
					}

					TD.ReportRowLeft
					{
					height: 26px;
					background-color: #F0F0F0;
					padding-left: 6px;
					width: 500px;
					}
					TD.ReportRow
					{
					width: 100px;
					border-left: 1px solid black;
					text-align: left;
					background-color: #F0F0F0;
					padding-left: 6px;
					}
					TD.ReportRowHeader
					{
					border-bottom: 1px solid black;
					border-left: 1px solid black;
					height: 60px;
					background-color: #C0C0C0;
					}
					Table.ReportHeader
					{
					border:1px solid black;
					border-collapse: collapse;
					font-family:calibri;
					font-size:14px;
					}

				</style>
				
			</head>
			
			<body style="font-family:calibri; font-size:12px;">
				<center>
					<table style="width:400px;" class="ReportHeader">
						<tr>
							<td style="width:200px; height:30px; padding-left:10px; padding-top:10px;">
								MIME Types from IIS					
							</td>
							<td style="width:200px; padding-right:10px; padding-top:10px; text-align:right;">
								Date : <xsl:value-of select="$Data" />
							</td>
						</tr>
					</table>
					<br />
					<table style="width:400px;" class="Report">
						<tr>
							<td style="width:200px; padding-left: 6px;" class="ReportRowHeader">
								MIME Type
							</td>
							<td style="width: 200px; text-align: center;" class="ReportRowHeader">
								Extension
							</td>
						</tr>
						<xsl:for-each select="RegisteredType">
							<tr>
								<td class="ReportRowLeft">
									<xsl:value-of select="MimeType"/>
								</td>
								<td class="ReportRow">
									<xsl:value-of select="Extension"/>
								</td>
							</tr>
						</xsl:for-each>
					</table>
					<br />
					<table style="width:400px;" class="ReportHeader">
						<tr>
							<td style="width:200px; height:40px; padding-left:10px;">
						
							</td>
							<td style="width:200px; padding-right:10px; padding-top:10px; text-align:right;">
								Date : <xsl:value-of select="$Data" />
							</td>
						</tr>
					</table>
					<br />
				</center>
			</body>
		</html>
	</xsl:template>
</xsl:stylesheet>