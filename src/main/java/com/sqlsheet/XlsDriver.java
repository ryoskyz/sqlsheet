/*
 * Copyright 2012 pcal.net
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

import com.sqlsheet.stream.XlsStreamConnection;

import org.apache.commons.vfs2.FileObject;
import org.apache.commons.vfs2.VFS;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.InputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.DriverPropertyInfo;
import java.sql.SQLException;
import java.sql.SQLFeatureNotSupportedException;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * SqlSheet implementation of java.sql.Driver.
 *
 * @author <a href="http://www.pcal.net">pcal</a>
 * @author <a href="http://code.google.com/p/sqlsheet">sqlsheet</a>
 */
public class XlsDriver implements java.sql.Driver {

    public static final String READ_STREAMING = "readStreaming";
    public static final String WRITE_STREAMING = "writeStreaming";
    public static final String HEADLINE = "headLine";
    public static final String FIRST_COL = "firstColumn";
    public static final String URL_SCHEME = "jdbc:xls:";
    public static final Logger LOGGER = Logger.getLogger(XlsDriver.class.getName());
    private static final Pattern CLASSPATH_OR_RESOURCE_PATTERN =
            Pattern.compile("^(classpath|resource):", Pattern.CASE_INSENSITIVE);
    private static final Pattern XLSX_PATTERN =
            Pattern.compile("^xlsx$", Pattern.CASE_INSENSITIVE);

    static {
        try {
            DriverManager.registerDriver(new XlsDriver());
        } catch (SQLException e) {
            LOGGER.log(Level.SEVERE, "Couldn't register " + XlsDriver.class.getName(), e);
        }
    }

    /**
     * @return the actual $HOME folder of the user
     */
    public static File getHomeFolder() {
        return new File(System.getProperty("user.home"));
    }

    /**
     * @param uriStr the String representation of an URI containing "~" or "${user.home}"
     * @return the expanded URI (resolving "~" and "${user.home}" to the actual $HOME folder
     */
    public static String resolveHomeUriStr(String uriStr) {
        String homePathStr = getHomeFolder().toURI().getPath();

        String expandedURIStr = uriStr.replaceFirst("~", Matcher.quoteReplacement(homePathStr));
        expandedURIStr =
                expandedURIStr.replaceFirst("\\$\\{user.home}",
                        Matcher.quoteReplacement(homePathStr));

        return expandedURIStr;
    }

    @Override
    public DriverPropertyInfo[] getPropertyInfo(String url, Properties info) {
        return new DriverPropertyInfo[0];
    }

    /**
     * Attempts to make a database connection to the given URL. The driver should return "null" if
     * it realizes it is the wrong kind of driver to connect to the given URL. This will be common,
     * as when the JDBC driver manager is asked to connect to a given URL it passes the URL to each
     * loaded driver in turn.
     *
     * <p>
     * The driver should throw an <code>SQLException</code> if it is the right driver to connect to
     * the given URL but has trouble connecting to the database.
     *
     * <p>
     * The {@code url} should point to supported file systems. (e.g., a file or a resource in the
     * class path)
     *
     * <p>
     * Valid samples are:
     *
     * <ul>
     * <li>jdbc:xls:file://${user.home}/dataSource.xlsx
     * <li>jdbc:xls:file://~/dataSource.xlsx
     * <li>jdbc:xls:classpath:/com/sqlsheet/dataSource.xlsx
     * <li>jdbc:xls:resource:/com/sqlsheet/dataSource.xlsx
     * </ul>
     *
     * <p>
     * The {@code Properties} argument can be used to pass arbitrary string tag/value pairs as
     * connection arguments. Normally at least "user" and "password" properties should be included
     * in the {@code Properties} object.
     *
     * <p>
     * <B>Note:</B> If a property is specified as part of the {@code url} and is also specified in
     * the {@code Properties} object, it is implementation-defined as to which value will take
     * precedence. For maximum portability, an application should only specify a property once.
     *
     * @param url the URL of the database to which to connect
     * @param info a list of arbitrary string tag/value pairs as connection arguments. Normally at
     *        least a "user" and "password" property should be included.
     * @return a <code>Connection</code> object that represents a connection to the URL
     * @throws SQLException if a database access error occurs or the url is {@code null}
     * @see <a href="https://commons.apache.org/proper/commons-vfs/filesystems.html"> Supported File
     *      Systems(Apache Commons VFS)</a>
     */
    @SuppressWarnings("PMD.NPathComplexity")
    public Connection connect(String url, Properties info) throws SQLException {
        if (url == null) {
            throw new IllegalArgumentException("Null url");
        }
        if (!acceptsURL(url)) {
            return null; // why is this necessary?
        }
        if (!url.toLowerCase().startsWith(URL_SCHEME)) {
            throw new IllegalArgumentException("URL is not " + URL_SCHEME + " (" + url + ")");
        }
        // strip any properties from end of URL and set them as additional properties
        String urlProperties;
        int questionIndex = url.indexOf('?');
        if (questionIndex >= 0) {
            urlProperties = url.substring(questionIndex);
            String[] split = urlProperties.substring(1).split("&");
            for (String each : split) {
                String[] property = each.split("=");
                try {
                    if (property.length == 2) {
                        String key = URLDecoder.decode(property[0], "UTF-8");
                        String value = URLDecoder.decode(property[1], "UTF-8");
                        info.setProperty(key, value);
                    } else if (property.length == 1) {
                        String key = URLDecoder.decode(property[0], "UTF-8");
                        info.setProperty(key, Boolean.TRUE.toString());
                    } else {
                        throw new SQLException("Invalid property: " + each);
                    }
                } catch (UnsupportedEncodingException e) {
                    // we know UTF-8 is available
                }
            }
        }
        String strippedUrlStr = questionIndex >= 0
                ? url.substring(0, questionIndex)
                : url;
        String workbookUriStr = strippedUrlStr.substring(URL_SCHEME.length());
        workbookUriStr = resolveHomeUriStr(workbookUriStr);
        workbookUriStr =
                CLASSPATH_OR_RESOURCE_PATTERN.matcher(workbookUriStr).replaceFirst("res:");
        try (FileObject file = VFS.getManager().resolveFile(workbookUriStr)) {
            // If streaming requested for read
            if (has(info, READ_STREAMING)) {
                return new XlsStreamConnection(file.getURL(), info);
            } else if (file.isWriteable()) {
                // If streaming requested for write
                boolean xlsx = XLSX_PATTERN.matcher(file.getName().getExtension()).matches();
                if (has(info, WRITE_STREAMING)) {
                    if (xlsx) {
                        return new XlsConnection(getOrCreateXlsxStream(file), file.getURL(), info);
                    }
                    LOGGER.warning(WRITE_STREAMING + " is not supported on " + strippedUrlStr);
                }
                return new XlsConnection(getOrCreateWorkbook(file, xlsx), file.getURL(), info);
            } else {
                try (InputStream in = file.getContent().getInputStream()) {
                    // If plain url provided
                    return new XlsConnection(WorkbookFactory.create(in), info);
                }
            }
        } catch (Exception e) {
            throw new SQLException(e.getMessage(), e);
        }
    }

    boolean has(Properties info, String key) {
        Object value = info.get(key);
        if (value == null) {
            return false;
        }
        return value.equals(Boolean.TRUE.toString());
    }

    private SXSSFWorkbook getOrCreateXlsxStream(FileObject file) throws IOException {
        if (!file.exists() && VFS.getManager().canCreateFileSystem(file)
                || file.getContent().getSize() == 0) {
            try (Workbook workbook = new XSSFWorkbook()) {
                flushWorkbook(workbook, file);
            }
        } else {
            LOGGER.warning(
                    "File " + file.getPath() + " is not empty, and will parsed to memory!");
        }
        try (InputStream in = file.getContent().getInputStream()) {
            return new SXSSFWorkbook(new XSSFWorkbook(in), 1000, false);
        }
    }

    private Workbook getOrCreateWorkbook(FileObject file, boolean xlsx) throws IOException {
        if (!file.exists() && VFS.getManager().canCreateFileSystem(file)
                || file.getContent().getSize() == 0) {
            try (Workbook workbook = xlsx ? new XSSFWorkbook() : new HSSFWorkbook()) {
                flushWorkbook(workbook, file);
            }
        }
        org.apache.poi.openxml4j.util.ZipInputStreamZipEntrySource
                .setThresholdBytesForTempFiles(100_000_000);
        IOUtils.setByteArrayMaxOverride(500_000_000);
        try (InputStream in = file.getContent().getInputStream()) {
            return WorkbookFactory.create(in);
        }
    }

    private void flushWorkbook(Workbook workbook, FileObject file) throws IOException {
        try (OutputStream fileOut = file.getContent().getOutputStream()) {
            workbook.write(fileOut);
            fileOut.flush();
        }
    }

    public boolean acceptsURL(String url) {
        return url != null && url.trim().toLowerCase().startsWith(URL_SCHEME);
    }

    public boolean jdbcCompliant() { // LOLZ!
        return true;
    }

    public int getMajorVersion() {
        return 1;
    }

    public int getMinorVersion() {
        return 0;
    }

    public Logger getParentLogger() throws SQLFeatureNotSupportedException {
        throw new SQLFeatureNotSupportedException("Not supported yet.");
    }
}
