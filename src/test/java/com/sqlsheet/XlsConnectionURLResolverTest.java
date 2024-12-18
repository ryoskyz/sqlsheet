/*
 * Copyright 2020 Andreas Reichel <andreas@manticore-projects.com>.
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
package com.sqlsheet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.DateFormatConverter;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.util.Locale;

/**
 * @author Andreas Reichel <andreas@manticore-projects.com>
 */
public class XlsConnectionURLResolverTest {

    public static final int DEFAULT_HEADLINE = 1;
    public static final int DEFAULT_FIRST_COL = 0;
    private final String[] columnNames = {"A", "B", "C"};
    private final CellType[] cellTypes = {CellType.NUMERIC, CellType.NUMERIC, CellType.FORMULA};
    private final Object[][] cellValues = {
            {1d, 2d, "A2+B2"},
            {3d, 4d, "A3+B3"},
            {5d, 6d, "A4+B4"}
    };
    private Connection conn;

    public XlsConnectionURLResolverTest() throws SQLException, IOException, ClassNotFoundException {
        Workbook workBook = WorkbookFactory.create(true);
        Sheet sheet = workBook.createSheet("TestSheet1");
        int r = DEFAULT_HEADLINE - 1;
        Row row = sheet.createRow(r);
        for (int c = DEFAULT_FIRST_COL; c < columnNames.length + DEFAULT_FIRST_COL; c++) {
            Cell cell = row.createCell(c);
            cell.setCellValue(columnNames[c - DEFAULT_FIRST_COL]);
        }

        DataFormat dateFormat = workBook.createDataFormat();
        short dateFormatIndex =
                dateFormat.getFormat(
                        DateFormatConverter.convert(
                                Locale.US,
                                DateFormat.getDateInstance(DateFormat.SHORT, Locale.US)));

        CellStyle dateCellStyle = workBook.createCellStyle();
        dateCellStyle.setDataFormat(dateFormatIndex);

        while (r < 3 + DEFAULT_HEADLINE - 1) {
            r++;
            row = sheet.createRow(r);
            for (int c = DEFAULT_FIRST_COL; c < columnNames.length + DEFAULT_FIRST_COL; c++) {
                Cell cell = row.createCell(c);
                switch (cellTypes[c - DEFAULT_FIRST_COL]) {
                    case NUMERIC:
                        cell.setCellValue(
                                (Double) cellValues[r - DEFAULT_HEADLINE][c - DEFAULT_FIRST_COL]);
                        break;
                    case FORMULA:
                        cell.setCellFormula(
                                (String) cellValues[r - DEFAULT_HEADLINE][c - DEFAULT_FIRST_COL]);
                        break;
                }
            }
        }

        File file = new File(XlsDriver.getHomeFolder(),
                XlsConnectionURLResolverTest.class.getSimpleName() + ".xlsx");
        if (file.exists() && file.canWrite()) {
            file.delete();
        } else {
            file.deleteOnExit();
        }

        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workBook.write(fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();

    }

    @BeforeAll
    public static void loadDriverClass() throws ClassNotFoundException {
        Class.forName("com.sqlsheet.XlsDriver");
    }

    @AfterEach
    public void tearDown() {
        try {
            if (conn != null && !conn.isClosed()) {
                conn.close();
            }
        } catch (SQLException ex) {
            // nothing we could do
        }
    }

    @Test
    public void connectionFromTildeHome() throws Exception {
        conn =
                DriverManager.getConnection(
                        "jdbc:xls:file://~/"
                                + XlsConnectionURLResolverTest.class.getSimpleName() + ".xlsx"
                                + "?headLine="
                                + DEFAULT_HEADLINE
                                + "&firstColumn="
                                + DEFAULT_FIRST_COL);

        Statement statement = null;
        ResultSet resultSet = null;
        try {
            statement = conn.createStatement();
            resultSet = statement.executeQuery("SELECT * FROM TestSheet1");
            Assertions.assertTrue(resultSet.next(), "Recordset should have records");

        } finally {
            if (resultSet != null && !resultSet.isClosed()) {
                try {
                    resultSet.close();
                } catch (SQLException ex1) {
                    // nothing we can do about
                }
            }

            if (statement != null && !statement.isClosed()) {
                try {
                    statement.close();
                } catch (SQLException ex1) {
                    // nothing we can do about
                }
            }

            if (conn != null && !conn.isClosed()) {
                try {
                    conn.close();
                } catch (SQLException ex1) {
                    // nothing we can do about
                }
            }
        }
    }

    @Test
    public void connectionFromSystemPropertyHome() throws Exception {
        conn =
                DriverManager.getConnection(
                        "jdbc:xls:file://${user.home}/"
                                + XlsConnectionURLResolverTest.class.getSimpleName() + ".xlsx"
                                + "?headLine="
                                + DEFAULT_HEADLINE
                                + "&firstColumn="
                                + DEFAULT_FIRST_COL);

        Statement statement = null;
        ResultSet resultSet = null;
        try {
            statement = conn.createStatement();
            resultSet = statement.executeQuery("SELECT * FROM TestSheet1");
            Assertions.assertTrue(resultSet.next(), "Recordset should have records");

        } finally {
            if (resultSet != null && !resultSet.isClosed()) {
                try {
                    resultSet.close();
                } catch (SQLException ex1) {
                    // nothing we can do about
                }
            }

            if (statement != null && !statement.isClosed()) {
                try {
                    statement.close();
                } catch (SQLException ex1) {
                    // nothing we can do about
                }
            }

            if (conn != null && !conn.isClosed()) {
                try {
                    conn.close();
                } catch (SQLException ex1) {
                    // nothing we can do about
                }
            }
        }
    }

    @Test
    public void connectionFromClassPathRessource() throws Exception {
        conn =
                DriverManager.getConnection(
                        "jdbc:xls:classpath:/headline.xlsx?headLine=5");

        Statement statement = null;
        ResultSet resultSet = null;
        try {
            statement = conn.createStatement();
            resultSet = statement.executeQuery("SELECT * FROM headline");
            Assertions.assertTrue(resultSet.next(), "Recordset should have records");

        } finally {
            if (resultSet != null && !resultSet.isClosed()) {
                try {
                    resultSet.close();
                } catch (SQLException ex1) {
                    // nothing we can do about
                }
            }

            if (statement != null && !statement.isClosed()) {
                try {
                    statement.close();
                } catch (SQLException ex1) {
                    // nothing we can do about
                }
            }

            if (conn != null && !conn.isClosed()) {
                try {
                    conn.close();
                } catch (SQLException ex1) {
                    // nothing we can do about
                }
            }
        }
    }
}
