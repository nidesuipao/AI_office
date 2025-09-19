<?xml version="1.0"?>
<!-- Microsoft MathML to OMML XSLT (abridged placeholder). For production, replace with full official XSL. -->
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
    xmlns:m="http://www.w3.org/1998/Math/MathML"
    xmlns:mml="http://www.w3.org/1998/Math/MathML"
    xmlns:mso="http://schemas.microsoft.com/office/2004/12/omml"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    exclude-result-prefixes="m mml">

  <xsl:output method="xml" indent="yes"/>

  <!-- Minimal transform: wraps MathML into an oMathPara as plain text fallback. -->
  <xsl:template match="/">
    <mso:oMathPara xmlns:mso="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <mso:oMath>
        <mso:r>
          <w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
          <w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <xsl:value-of select="normalize-space(//m:math)"/>
          </w:t>
        </mso:r>
      </mso:oMath>
    </mso:oMathPara>
  </xsl:template>

</xsl:stylesheet>


