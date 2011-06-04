<xsl:stylesheet version="1.0"
    xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

    <xsl:variable name="block-elements" select="'h1 h2 h3 h4 h5 h6 p div table ul ol pre img'" />
    <xsl:variable name="inline-elements" select="'strong em span'" />

    <xsl:template match="head|script|style"/>

    <xsl:template match="body">
        <w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
            <w:body>
                <xsl:for-each select="node()">
                    <xsl:variable name="preceding-element" select="local-name(preceding-sibling::*[position()=1])"/>
                    <!-- test if we are preceded by a block element -->
                    <xsl:if test="contains(concat(' ', $block-elements, ' '),
                                        concat(' ', $preceding-element, ' '))">
                        <xsl:if test="contains(concat(' ', $inline-elements, ' '),
                                            concat(' ', name(), ' ')) or
                                    name()='' and normalize-space(.) != ''">
                            <w:p>
                                <w:pPr/>
                                <xsl:apply-templates select="." mode="group-inline-siblings"/>
                            </w:p>
                        </xsl:if>
                    </xsl:if>
                    <xsl:if test="contains(concat(' ', $block-elements, ' '),
                                        concat(' ', name(), ' '))">
                        <xsl:apply-templates select="." mode="block-elements"/>
                    </xsl:if>
                </xsl:for-each>
            </w:body>
        </w:document>
    </xsl:template>

    <!-- template to handle nested block elements -->
    <xsl:template match="h1|h2|h3|h4|h5|h6|p|div|ul|ol|pre|img" mode="block-elements">
        <!-- recurse if we contain other block elements -->
        <xsl:if test="h1|h2|h3|h4|h5|h6|p|div|ul|ol|pre">
            <xsl:for-each select="node()">
                <xsl:apply-templates select="." mode="block-elements"/>
            </xsl:for-each>
        </xsl:if>
        <xsl:if test="not(h1|h2|h3|h4|h5|h6|p|div|ul|ol|pre)">
            <xsl:apply-templates select="."/>
        </xsl:if>
    </xsl:template>

    <xsl:template match="node()" mode="group-inline-siblings">
        <xsl:variable name="name" select="local-name(.)"/>
        <!-- check if we are a block element -->
        <xsl:if test="not(contains(concat(' ', $block-elements, ' '),
                                   concat(' ', $name, ' ')))">
            <xsl:apply-templates select="."/>
            <xsl:apply-templates select="following-sibling::node()[position()=1]" mode="group-inline-siblings"/>
        </xsl:if>
    </xsl:template>

    <xsl:template match="em" mode="ancestor_properties">
        <w:i/>
    </xsl:template>

    <xsl:template match="strong" mode="ancestor_properties">
        <w:b/>
    </xsl:template>

    <xsl:template match="pre" mode="ancestor_properties">
        <w:rStyle w:val="sourcetext"/>
    </xsl:template>

    <xsl:template match="div" mode="ancestor_properties">
        <w:rStyle w:val="quotation"/>
    </xsl:template>


    <xsl:template match="text()" >
        <xsl:if test="normalize-space(.) != ''">
            <w:r>
                <w:rPr>
                    <xsl:apply-templates select="ancestor::em|ancestor::strong|ancestor::pre|ancestor::div[@class='pullquote']" mode="ancestor_properties"/>
                </w:rPr>
                <w:t><xsl:value-of select="."/><xsl:text> </xsl:text></w:t>
            </w:r>
        </xsl:if>
    </xsl:template>

    <xsl:template match="p">
        <w:p>
            <w:pPr/>
            <xsl:apply-templates/>
        </w:p>
    </xsl:template>

    <xsl:template match="h2">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="heading2"/>
            </w:pPr>
            <xsl:apply-templates/>
        </w:p>
    </xsl:template>

    <xsl:template match="h3">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="heading3"/>
            </w:pPr>
            <xsl:apply-templates/>
        </w:p>
    </xsl:template>

    <xsl:template match="pre">
        <w:p>
            <w:pPr/>
            <xsl:apply-templates/>
        </w:p>
    </xsl:template>

    <xsl:template match="div">
        <w:p>
            <w:pPr>
                <xsl:if test="@align='left'">
                    <w:jc w:val="left"/>
                </xsl:if>
                <xsl:if test="@align='center'">
                    <w:jc w:val="center"/>
                </xsl:if>
                <xsl:if test="@align='right'">
                    <w:jc w:val="right"/>
                </xsl:if>
            </w:pPr>
            <xsl:apply-templates/>
        </w:p>
    </xsl:template>

    <xsl:template match="ol/li">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="textbody"/>
                <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="4"/>
                </w:numPr>
            </w:pPr>
            <xsl:apply-templates/>
        </w:p>
    </xsl:template>

    <xsl:template match="ul/li">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="textbody"/>
                <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="5"/>
                </w:numPr>
            </w:pPr>
            <xsl:apply-templates/>
        </w:p>
    </xsl:template>

    <xsl:template match="table" mode="block-elements">
        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="0" w:type="auto"/>
                <!-- we include borders by default -->
                <w:tblBorders>
                    <w:top w:color="auto" w:space="0" w:sz="2" w:val="single"/>
                    <w:left w:color="auto" w:space="0" w:sz="2" w:val="single"/>
                    <w:bottom w:color="auto" w:space="0" w:sz="2" w:val="single"/>
                    <w:right w:color="auto" w:space="0" w:sz="2" w:val="single"/>
                    <w:insideH w:color="auto" w:space="0" w:sz="2" w:val="single"/>
                    <w:insideV w:color="auto" w:space="0" w:sz="2" w:val="single"/>
                </w:tblBorders>
                <w:jc w:val="left"/>
            </w:tblPr>
            <w:tblGrid>
                <xsl:variable name="maxCells">
                    <xsl:for-each select="*/tr">
                        <xsl:sort select="count(td[not(@colspan)]) + sum(td/@colspan)" order="descending"/>
                            <xsl:if test="position() = 1">
                                <xsl:value-of select="count(td[not(@colspan)]) + sum(td/@colspan)"/>
                            </xsl:if>
                    </xsl:for-each>
                </xsl:variable>
                <xsl:call-template name="gridColumns">
                    <xsl:with-param name="numColumns" select="$maxCells"/>
                </xsl:call-template>
            </w:tblGrid>
            <xsl:for-each select="*/tr">
                <xsl:call-template name="tr"/>
            </xsl:for-each>
        </w:tbl>
    </xsl:template>

    <xsl:template name="gridColumns">
        <xsl:param name="numColumns" select="0"/>

        <xsl:if test="$numColumns > 0">
            <w:gridCol w:w="1200"/>

            <xsl:call-template name="gridColumns">
                <xsl:with-param name="numColumns" select="$numColumns -1"/>
            </xsl:call-template>
        </xsl:if>
    </xsl:template>

    <xsl:template name="tr">
        <w:tr>
            <w:trPr>
                <xsl:if test="ancestor::thead">
                    <w:tblHeader w:val="true"/>
                </xsl:if>
                <w:cantSplit w:val="false"/>
            </w:trPr>
            <xsl:for-each select="th">
                <xsl:call-template name="cell"/>
            </xsl:for-each>
            <xsl:for-each select="td">
                <xsl:call-template name="cell"/>
            </xsl:for-each>
        </w:tr>
    </xsl:template>

    <xsl:template name="cell">
        <w:tc>
            <w:p>
                <w:pPr>
                    <xsl:if test="name() = 'th'">
                        <w:pStyle w:val="tableheading"/>
                    </xsl:if>
                    <xsl:if test="name() = 'td'">
                        <w:pStyle w:val="tablecontents"/>
                    </xsl:if>
                </w:pPr>
                <xsl:apply-templates />
            </w:p>
        </w:tc>
    </xsl:template>

    <xsl:template match="img">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="textbody"/>
                <w:jc w:val="left"/>
                <w:spacing w:after="120" w:before="0"/>
            </w:pPr>
            <w:r>
                <w:rPr/>
                <w:drawing>
                    <wp:inline>
                        <!-- width and height will be computed and transformed in python -->
                        <wp:extent>
                            <xsl:attribute name="cx">
                                <xsl:value-of select="./@src" />
                                <xsl:text>-$width</xsl:text>
                            </xsl:attribute>
                            <xsl:attribute name="cy">
                                <xsl:value-of select="./@src" />
                                <xsl:text>-$height</xsl:text>
                            </xsl:attribute>
                        </wp:extent>
                        <wp:docPr>
                            <xsl:attribute name="name">
                                <xsl:value-of select="./@src" />
                            </xsl:attribute>
                            <xsl:attribute name="id">
                                <xsl:value-of select="./@src" />
                            </xsl:attribute>
                        </wp:docPr>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/3/main">
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/3/picture">
                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/3/picture">
                            <pic:nvPicPr>
                                <pic:cNvPr>
                                    <xsl:attribute name="name">
                                        <xsl:value-of select="./@src" />
                                    </xsl:attribute>
                                    <xsl:attribute name="id">
                                        <xsl:value-of select="./@src" />
                                    </xsl:attribute>
                                </pic:cNvPr>
                                <pic:cNvPicPr />
                            </pic:nvPicPr>
                            <pic:blipFill>
                                <a:blip r:embed="R564dd6f276f347c3" />
                                <a:blip>
                                    <xsl:attribute name="r:embed">
                                        <xsl:value-of select="./@src" />
                                        <xsl:text>-$rid</xsl:text>
                                    </xsl:attribute>
                                </a:blip>
                                <a:stretch>
                                <a:fillRect />
                                </a:stretch>
                            </pic:blipFill>
                            <pic:spPr>
                                <a:xfrm>
                                    <a:off x="0" y="0" />
                                    <a:ext>
                                        <xsl:attribute name="cx">
                                            <xsl:value-of select="./@src" />
                                            <xsl:text>-$width</xsl:text>
                                        </xsl:attribute>
                                        <xsl:attribute name="cy">
                                            <xsl:value-of select="./@src" />
                                            <xsl:text>-$height</xsl:text>
                                        </xsl:attribute>
                                    </a:ext>
                                </a:xfrm>
                                <a:prstGeom prst="rect" />
                            </pic:spPr>
                            </pic:pic>
                        </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
    </xsl:template>

</xsl:stylesheet>

