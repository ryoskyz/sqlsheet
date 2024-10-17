package com.sqlsheet.stream;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.util.List;


public class XlsxSheetIteratorTest {

    @Test
    public void testMain() throws Exception {
        XlsxSheetIterator xlsSheet =
                new XlsxSheetIterator(ClassLoader.getSystemResource("test.xlsx"), "2009");
        Assertions.assertEquals(3, xlsSheet.getColumns().size());
        long counter = 0L;
        for (List<XlsSheetIterator.CellValueHolder> rowValues : xlsSheet) {
            Assertions.assertEquals(3, rowValues.size());
            for (XlsSheetIterator.CellValueHolder value : rowValues) {
                Assertions.assertNotNull(value.stringValue);
            }
            Assertions.assertNotNull(rowValues.get(0).doubleValue);
            Assertions.assertNotNull(rowValues.get(2).dateValue);
            counter++;
        }
        Assertions.assertEquals(3L, counter);
    }

    @Test
    public void testBigGrid() throws Exception {
        XlsxSheetIterator xlsSheet =
                new XlsxSheetIterator(ClassLoader.getSystemResource("big-grid.xlsx"), "Big Grid");
        Assertions.assertEquals(20, xlsSheet.getColumns().size());
        long counter = 0L;
        for (List<XlsSheetIterator.CellValueHolder> rowValues : xlsSheet) {
            Assertions.assertEquals(20, rowValues.size());
            for (XlsSheetIterator.CellValueHolder value : rowValues) {
                Assertions.assertNotNull(value.stringValue);
            }
            Assertions.assertNotNull(rowValues.get(1).doubleValue);
            Assertions.assertNotNull(rowValues.get(2).doubleValue);
            Assertions.assertNotNull(rowValues.get(3).doubleValue);
            Assertions.assertNotNull(rowValues.get(4).dateValue);
            counter++;
        }
        Assertions.assertEquals(65535L, counter);
    }

    @Test
    public void testBugNo3() throws Exception {
        XlsxSheetIterator xlsSheet =
                new XlsxSheetIterator(ClassLoader.getSystemResource("bug3.xlsx"), "bug3");
        Assertions.assertEquals(13, xlsSheet.getColumns().size());
        long counter = 0L;
        for (List<XlsSheetIterator.CellValueHolder> rowValues : xlsSheet) {
            Assertions.assertEquals(13, rowValues.size());
            // for(XlsSheetIterator.CellValueHolder value : rowValues){
            // Assert.assertNotNull(value.stringValue);
            // }
            counter++;
        }
        Assertions.assertEquals(1, counter);
    }

}
