/**
 *
 * @author Wei-Ming Wu
 *
 *
 * Copyright 2013 Wei-Ming Wu
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not
 * use this file except in compliance with the License. You may obtain a copy of
 * the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
 * License for the specific language governing permissions and limitations under
 * the License.
 *
 */
package com.github.wnameless.workbookaccessor;

import static net.sf.rubycollect4j.RubyCollections.Hash;
import static net.sf.rubycollect4j.RubyCollections.newRubyArray;
import static net.sf.rubycollect4j.RubyCollections.newRubyLazyEnumerator;
import static net.sf.rubycollect4j.RubyCollections.ra;
import static net.sf.rubycollect4j.RubyCollections.range;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import net.sf.rubycollect4j.RubyArray;
import net.sf.rubycollect4j.RubyLazyEnumerator;
import net.sf.rubycollect4j.block.TransformBlock;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.github.wnameless.nullproof.annotation.AcceptNull;
import com.github.wnameless.nullproof.annotation.RejectNull;

/**
 * 
 * WorkbookReader is a wrapper to Apache POI. It tends to provide friendly APIs
 * for workbook reading.
 * 
 */
@RejectNull
public final class WorkbookReader {

  private static final Logger logger = Logger.getLogger(WorkbookReader.class
      .getName());

  private NPOIFSFileSystem npoifs = null;
  private OPCPackage pkg = null;
  private Workbook wb;
  private Sheet sheet;
  private final RubyArray<String> header = newRubyArray();
  private boolean hasHeader = true;
  private boolean isClosed = false;

  /**
   * Creates a WorkbookReader by given path. Assumes there is a header within
   * the spreadsheet.
   * 
   * @param path
   *          of a Workbook
   */
  public WorkbookReader(String path) {
    File file = new File(path);
    setSourceFile(file);
    sheet = wb.getSheetAt(0);
    setHeader();
  }

  /**
   * Creates a WorkbookReader by given path.
   * 
   * @param path
   *          of a Workbook
   * @param hasHeader
   *          true if spreadsheet gets a header, false otherwise
   */
  public WorkbookReader(String path, boolean hasHeader) {
    File file = new File(path);
    setSourceFile(file);
    sheet = wb.getSheetAt(0);
    this.hasHeader = hasHeader;
    setHeader();
  }

  /**
   * Creates a WorkbookReader by given File. Assumes there is a header within
   * the spreadsheet.
   * 
   * @param file
   *          of a Workbook
   */
  public WorkbookReader(File file) {
    setSourceFile(file);
    sheet = wb.getSheetAt(0);
    setHeader();
  }

  /**
   * Creates a WorkbookReader by given File.
   * 
   * @param file
   *          of a Workbook
   * @param hasHeader
   *          true if spreadsheet gets a header, false otherwise
   */
  public WorkbookReader(File file, boolean hasHeader) {
    setSourceFile(file);
    sheet = wb.getSheetAt(0);
    this.hasHeader = hasHeader;
    setHeader();
  }

  /**
   * Creates a WorkbookReader by given Workbook. Assumes there is a header
   * within the spreadsheet.
   * 
   * @param wb
   *          a Workbook
   */
  public WorkbookReader(Workbook wb) {
    this.wb = wb;
    sheet = wb.getSheetAt(0);
    hasHeader = true;
    setHeader();
  }

  /**
   * Creates a WorkbookReader by given Workbook.
   * 
   * @param wb
   *          a Workbook
   * @param hasHeader
   *          true if spreadsheet gets a header, false otherwise
   */
  public WorkbookReader(Workbook wb, boolean hasHeader) {
    this.wb = wb;
    sheet = wb.getSheetAt(0);
    this.hasHeader = hasHeader;
    setHeader();
  }

  private void setSourceFile(File file) {
    try {
      npoifs = new NPOIFSFileSystem(file);
      wb = WorkbookFactory.create(npoifs);
    } catch (OfficeXmlFileException ofe) {
      try {
        pkg = OPCPackage.open(file);
        wb = WorkbookFactory.create(pkg);
      } catch (Exception e) {
        logger.log(Level.SEVERE, null, e);
        throw new RuntimeException(e);
      }
    } catch (IOException e) {
      logger.log(Level.SEVERE, null, e);
      throw new RuntimeException(e);
    }
  }

  /**
   * Manually closes the Workbook file.
   */
  public void close() {
    if (npoifs != null) {
      try {
        npoifs.close();
      } catch (IOException e) {
        logger.log(Level.SEVERE, null, e);
        throw new RuntimeException(e);
      }
    }
    if (pkg != null) {
      try {
        pkg.close();
      } catch (IOException e) {
        logger.log(Level.SEVERE, null, e);
        throw new RuntimeException(e);
      }
    }
    isClosed = true;
  }

  /**
   * Returns the backing POI Workbook.
   * 
   * @return the POI Workbook
   */
  public Workbook getWorkbook() {
    return wb;
  }

  private void setHeader() {
    header.clear();
    Iterator<Row> rows = sheet.rowIterator();
    if (rows.hasNext() && hasHeader)
      header.concat(rowToRubyArray(rows.next()));
  }

  /**
   * Returns a List which contains all header fields.
   * 
   * @return a String List
   */
  public List<String> getHeader() {
    if (isClosed)
      throw new IllegalStateException("Workbook has been closed.");

    return header.each().toA();
  }

  /**
   * Returns the name of current sheet.
   * 
   * @return the sheet name
   */
  public String getCurrentSheetName() {
    return sheet.getSheetName();
  }

  /**
   * Returns a List which contains all sheet names.
   * 
   * @return a String List
   */
  public List<String> getAllSheetNames() {
    if (isClosed)
      throw new IllegalStateException("Workbook has been closed.");

    List<String> sheets = newRubyArray();
    for (int i = 0; i < wb.getNumberOfSheets(); i++) {
      sheets.add(wb.getSheetName(i));
    }
    return sheets;
  }

  /**
   * Turns this WorkbookReader to certain sheet. Sheets can be found by
   * getAllSheetNames().
   * 
   * @param index
   *          of a sheet
   * @return this WorkbookReader
   */
  public WorkbookReader turnToSheet(int index) {
    if (isClosed)
      throw new IllegalStateException("Workbook has been closed.");

    sheet = wb.getSheetAt(index);
    setHeader();
    return this;
  }

  /**
   * Turns this WorkbookReader to certain sheet. Sheets can be found by
   * getAllSheetNames().
   * 
   * @param sheetName
   *          name of a sheet
   * @return this WorkbookReader
   */
  public WorkbookReader turnToSheet(String sheetName) {
    if (!getAllSheetNames().contains(sheetName))
      throw new IllegalArgumentException("Sheet name is not found.");

    return turnToSheet(getAllSheetNames().indexOf(sheetName));
  }

  /**
   * Turns this WorkbookReader to certain sheet. Sheets can be found by
   * getAllSheetNames().
   * 
   * @param index
   *          of a sheet
   * @param hasHeader
   *          true if spreadsheet gets a header, false otherwise
   * @return this WorkbookReader
   */
  public WorkbookReader turnToSheet(int index, boolean hasHeader) {
    if (isClosed)
      throw new IllegalStateException("Workbook has been closed.");

    this.hasHeader = hasHeader;
    sheet = wb.getSheetAt(index);
    setHeader();
    return this;
  }

  /**
   * Turns this WorkbookReader to certain sheet. Sheets can be found by
   * getSheets().
   * 
   * @param sheetName
   *          name of a sheet
   * @param hasHeader
   *          true if spreadsheet gets a header, false otherwise
   * @return this WorkbookReader
   */
  public WorkbookReader turnToSheet(String sheetName, boolean hasHeader) {
    if (!getAllSheetNames().contains(sheetName))
      throw new IllegalArgumentException("Sheet name is not found.");

    return turnToSheet(getAllSheetNames().indexOf(sheetName), hasHeader);
  }

  /**
   * Converts the spreadsheet to CSV by a String Iterable.
   * 
   * @return a String Iterable
   */
  public Iterable<String> toCSV() {
    if (isClosed)
      throw new IllegalStateException("Workbook has been closed.");

    RubyLazyEnumerator<String> CSVIterable =
        newRubyLazyEnumerator(sheet).map(new TransformBlock<Row, String>() {

          @Override
          public String yield(Row item) {
            return rowToRubyArray(item, true).join(",");
          }

        });

    return hasHeader ? CSVIterable.drop(1) : CSVIterable;
  }

  /**
   * Converts the spreadsheet to String Lists by a List Iterable.
   * 
   * @return a String List Iterable
   */
  public Iterable<List<String>> toLists() {
    if (isClosed)
      throw new IllegalStateException("Workbook has been closed.");

    RubyLazyEnumerator<List<String>> listsIterable =
        newRubyLazyEnumerator(sheet).map(
            new TransformBlock<Row, List<String>>() {

              @Override
              public List<String> yield(Row item) {
                return rowToRubyArray(item);
              }

            });

    return hasHeader ? listsIterable.drop(1) : listsIterable;
  }

  /**
   * Converts the spreadsheet to String Arrays by an Array Iterable.
   * 
   * @return a String Array Iterable
   */
  public Iterable<String[]> toArrays() {
    if (isClosed)
      throw new IllegalStateException("Workbook has been closed.");

    RubyLazyEnumerator<String[]> arraysIterable =
        newRubyLazyEnumerator(sheet).map(new TransformBlock<Row, String[]>() {

          @Override
          public String[] yield(Row item) {
            List<String> list = rowToRubyArray(item);
            return list.toArray(new String[list.size()]);
          }

        });

    return hasHeader ? arraysIterable.drop(1) : arraysIterable;
  }

  /**
   * Converts the spreadsheet to Maps by a Map Iterable. All Maps are
   * implemented by LinkedHashMap which implies the order of all fields is kept.
   * 
   * @return a Map Iterable
   */
  public Iterable<Map<String, String>> toMaps() {
    if (isClosed)
      throw new IllegalStateException("Workbook has been closed.");
    if (!hasHeader)
      throw new IllegalStateException("Header is not found.");

    return newRubyLazyEnumerator(sheet).map(
        new TransformBlock<Row, Map<String, String>>() {

          @SuppressWarnings("unchecked")
          @Override
          public Map<String, String> yield(Row item) {
            return Hash(ra(getHeader()).zip(rowToRubyArray(item)));
          }

        }).drop(1);
  }

  private RubyArray<String> rowToRubyArray(final Row row) {
    return rowToRubyArray(row, false);
  }

  private RubyArray<String> rowToRubyArray(final Row row, final boolean isCSV) {
    RubyArray<Cell> cells;
    if (hasHeader) {
      int colNum = ra(sheet.rowIterator().next()).count();
      cells = range(0, colNum - 1).map(new TransformBlock<Integer, Cell>() {

        public Cell yield(Integer item) {
          return row.getCell(item);
        }

      });
    } else {
      cells = ra(row.cellIterator());
    }

    return cells.map(new TransformBlock<Cell, String>() {

      @AcceptNull
      public String yield(Cell item) {
        if (item == null)
          return "";

        item.setCellType(Cell.CELL_TYPE_STRING);
        String val = item.toString();
        if (isCSV && val.contains(",")) {
          val = val.replaceAll("\"", "\"\"");
          return '"' + val + '"';
        }
        return val;
      }

    });
  }

}
