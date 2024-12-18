/*
 * Copyright 2012 sqlsheet.googlecode.com
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License. You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under the License
 * is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied. See the License for the specific language governing permissions and limitations under
 * the License.
 */
package com.sqlsheet.stream;

import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.Characters;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;
import java.net.URL;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * Streaming iterator over XLSX files Derived from:
 * http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xssf/eventusermodel/XLSX2CSV.java
 */
public class XlsxSheetIterator extends AbstractXlsSheetIterator {
    private static final MathContext CTX_NN_15_EVEN = new MathContext(15, RoundingMode.HALF_EVEN);

    OPCPackage xlsxPackage;
    InputStream stream;
    XMLEventReader reader;
    StylesTable styles;
    ReadOnlySharedStringsTable strings;
    XSSFSheetEventHandler handler;

    public XlsxSheetIterator(URL filename, String sheetName) throws SQLException {
        super(filename, sheetName);
    }

    @Override
    protected void postConstruct() throws SQLException {
        try {
            // Open and pre process XLSX file
            xlsxPackage = OPCPackage.open(getFileName().getPath(), PackageAccess.READ);
            strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
            XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
            styles = xssfReader.getStylesTable();
            // Find appropriate sheet
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            while (iter.hasNext()) {
                stream = iter.next();
                String currentSheetName = iter.getSheetName();
                String quotedSheetName = "\"" + currentSheetName + "\"";
                if (currentSheetName.equalsIgnoreCase(getSheetName())
                        || quotedSheetName.equalsIgnoreCase(getSheetName())) {
                    handler = new XSSFSheetEventHandler(styles, strings);
                    XMLInputFactory factory = XMLInputFactory.newInstance();
                    reader = factory.createXMLEventReader(stream);
                    // Start sheet processing
                    while (reader.hasNext() && getCurrentSheetRowIndex() == 0) {
                        processNextEvent();
                    }
                    processNextRecords();
                } else {
                    IOUtils.closeQuietly(stream);
                }
            }
        } catch (Exception e) {
            throw new SQLException(e.getMessage(), e);
        }
    }

    @Override
    protected void processNextRecords() throws SQLException {
        Long nextRowIndex = getCurrentSheetRowIndex() + 2L;
        while (reader.hasNext() && !getCurrentSheetRowIndex().equals(nextRowIndex)) {
            try {
                processNextEvent();
            } catch (XMLStreamException e) {
                throw new SQLException(e.getMessage(), e);
            }
        }
    }

    @Override
    protected void onClose() {
        try {
            if (reader != null) {
                reader.close();
            }
        } catch (XMLStreamException e) {
            // not much we can do here
        } finally {
            IOUtils.closeQuietly(stream);
            IOUtils.closeQuietly(xlsxPackage);
        }
    }

    /**
     * Parses and shows the content of one sheet using the specified styles and shared-strings
     * tables.
     *
     * @throws XMLStreamException if any
     */
    public void processNextEvent() throws XMLStreamException {
        if (reader.hasNext()) {
            XMLEvent event = reader.nextEvent();
            XMLEvent nextEvent = reader.peek();
            switch (event.getEventType()) {
                case XMLEvent.START_ELEMENT:
                    handler.startElement(event.asStartElement());
                    if (nextEvent.isCharacters()) {
                        Characters c = reader.nextEvent().asCharacters();
                        if (!c.isWhiteSpace()) {
                            handler.characters(c.getData().toCharArray());
                        }
                    }
                    break;
                case XMLEvent.END_ELEMENT:
                    handler.endElement(event.asEndElement());
                    break;
                default:
                    // nothing
            }
        }
    }

    /**
     * The type of the data value is indicated by an attribute on the cell. The value is usually in
     * a "v" element within the cell.
     */
    enum XssfDataType {
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER,
    }

    /**
     * Derived from <a href="http://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api">...</a>
     *
     * <p>
     * Also see Standard ECMA-376, 1st edition, part 4, pages 1928ff, at
     * <a href="http://www.ecma-international.org/publications/standards/Ecma-376.htm">...</a>
     *
     * <p>
     */
    class XSSFSheetEventHandler {

        private final DataFormatter formatter;
        /**
         * Table with styles
         */
        private final StylesTable stylesTable;
        /**
         * Table with unique strings
         */
        private final ReadOnlySharedStringsTable sharedStringsTable;
        // Gathers characters as they are seen.
        private final StringBuffer value;
        // Set when V start element is seen
        private boolean vIsOpen;
        // Set when cell start element is seen;
        // used when cell close element is seen.
        private XssfDataType nextDataType;
        // Used to format numeric cell values.
        private short formatIndex;
        private String formatString;
        private int thisColumn;
        // The last column printed to the output stream
        private int lastColumnNumber;

        /**
         * Accepts objects needed while parsing.
         *
         * @param styles Table of styles
         * @param strings Table of shared strings
         */
        public XSSFSheetEventHandler(StylesTable styles, ReadOnlySharedStringsTable strings) {
            thisColumn = -1;
            lastColumnNumber = -1;
            this.stylesTable = styles;
            this.sharedStringsTable = strings;
            this.value = new StringBuffer();
            this.nextDataType = XssfDataType.NUMBER;
            this.formatter = new DataFormatter();
        }

        public void startElement(StartElement startElement) {
            Map<String, String> attributes = new HashMap<String, String>();
            Iterator<Attribute> attributesIterator = startElement.getAttributes();
            while (attributesIterator.hasNext()) {
                Attribute attr = attributesIterator.next();
                attributes.put(attr.getName().getLocalPart(), attr.getValue());
            }

            if ("inlineStr".equals(startElement.getName().getLocalPart())
                    || "v".equals(startElement.getName().getLocalPart())
                    || "is".equals(startElement.getName().getLocalPart())) {
                vIsOpen = true;
                // Clear contents cache
                value.setLength(0);
            } /* c => cell */ else if ("c".equals(startElement.getName().getLocalPart())) {
                // Get the cell reference
                String r = attributes.get("r");
                int firstDigit = -1;
                for (int c = 0; c < r.length(); ++c) {
                    if (Character.isDigit(r.charAt(c))) {
                        firstDigit = c;
                        break;
                    }
                }
                thisColumn = nameToColumn(r.substring(0, firstDigit));

                // Set up defaults.
                this.nextDataType = XssfDataType.NUMBER;
                this.formatIndex = -1;
                this.formatString = null;
                String cellType = attributes.get("t");
                String cellStyleStr = attributes.get("s");
                if ("b".equals(cellType)) {
                    nextDataType = XssfDataType.BOOL;
                } else if ("e".equals(cellType)) {
                    nextDataType = XssfDataType.ERROR;
                } else if ("inlineStr".equals(cellType)) {
                    nextDataType = XssfDataType.INLINESTR;
                } else if ("s".equals(cellType)) {
                    nextDataType = XssfDataType.SSTINDEX;
                } else if ("str".equals(cellType)) {
                    nextDataType = XssfDataType.FORMULA;
                } else if (cellStyleStr != null) {
                    // It's a number, but almost certainly one
                    // with a special style or format
                    int styleIndex = Integer.parseInt(cellStyleStr);
                    XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
                    this.formatIndex = style.getDataFormat();
                    this.formatString = style.getDataFormatString();
                    if (this.formatString == null) {
                        this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
                    }
                }
            }
        }

        public void endElement(EndElement endElement) {
            CellValueHolder thisCellValue = new CellValueHolder();
            // String thisStr = null;
            // v => contents of a cell
            if ("v".equals(endElement.getName().getLocalPart())
                    || "c".equals(endElement.getName().getLocalPart())
                            && XssfDataType.INLINESTR.equals(nextDataType)) {
                // Process the value contents as required.
                // Do now, as characters() may be called more than once
                switch (nextDataType) {
                    case BOOL:
                        char first = value.charAt(0);
                        thisCellValue.stringValue = first == '0' ? "FALSE" : "TRUE";
                        break;
                    case ERROR:
                        thisCellValue.stringValue = "\"ERROR:" + value + '"';
                        break;
                    case FORMULA:
                        // A formula could result in a string value,
                        // so always add double-quote characters.
                        thisCellValue.stringValue = value.toString();
                        break;
                    case INLINESTR:
                        // TODO: have seen an example of this, so it's untested.
                        XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
                        thisCellValue.stringValue = rtsi.toString();
                        break;
                    case SSTINDEX:
                        String sstIndex = value.toString();
                        try {
                            int idx = Integer.parseInt(sstIndex);

                            // @todo: check, if this is correct. I have never used RTF cell content
                            XSSFRichTextString rtss = new XSSFRichTextString(
                                    sharedStringsTable.getItemAt(idx).getString());
                            thisCellValue.stringValue = rtss.toString();
                        } catch (NumberFormatException ex) {
                            thisCellValue.stringValue =
                                    "Failed to parse SST index '" + sstIndex + "': "
                                            + ex;
                        }
                        break;
                    case NUMBER:
                        String n = value.toString();
                        final BigDecimal rawBigDecimal = new BigDecimal(n, CTX_NN_15_EVEN);
                        thisCellValue.doubleValue = rawBigDecimal.doubleValue();

                        if (this.formatString != null) {
                            thisCellValue.stringValue =
                                    formatter.formatRawCellContents(
                                            thisCellValue.doubleValue, this.formatIndex,
                                            this.formatString);
                        } else {
                            thisCellValue.stringValue = n;
                        }
                        thisCellValue.dateValue =
                                convertDateValue(thisCellValue.doubleValue, this.formatIndex,
                                        this.formatString);
                        break;
                    default:
                        thisCellValue.stringValue = "(TODO: Unexpected type: " + nextDataType + ")";
                        break;
                }
                // Output after we've seen the string contents
                // Emit commas for any fields that were missing on this row
                // Fill empty columns if required
                for (int i = lastColumnNumber + 1; i < thisColumn; ++i) {
                    // output.print(',');
                    if (getCurrentSheetRowIndex() == 0) {
                        getColumns().add(new CellValueHolder());
                    } else {
                        CellValueHolder empty = new CellValueHolder();
                        addCurrentRowValue(empty);
                    }
                }
                if (lastColumnNumber == -1) {
                    lastColumnNumber = 0;
                }
                // Might be the empty string.
                if (getCurrentSheetRowIndex() == 0) {
                    getColumns().add(thisCellValue);
                } else {
                    addCurrentRowValue(thisCellValue);
                }
                // Update column
                if (thisColumn > -1) {
                    lastColumnNumber = thisColumn;
                }
            } else if ("row".equals(endElement.getName().getLocalPart())) {
                // We're onto a new row
                lastColumnNumber = -1;
                setCurrentSheetRowIndex(getCurrentSheetRowIndex() + 1);
            }
        }

        /**
         * Captures characters only if a suitable element is open. Originally was just "v"; extended
         * for inlineStr also.
         */
        public void characters(char[] ch) {
            if (vIsOpen) {
                value.append(ch);
            }
        }

        /**
         * Converts an Excel column name like "C" to a zero-based index.
         *
         * @param name column name
         * @return Index corresponding to the specified name
         */
        private int nameToColumn(String name) {
            int column = -1;
            for (int i = 0; i < name.length(); ++i) {
                int c = name.charAt(i);
                column = (column + 1) * 26 + c - 'A';
            }
            return column;
        }
    }
}
