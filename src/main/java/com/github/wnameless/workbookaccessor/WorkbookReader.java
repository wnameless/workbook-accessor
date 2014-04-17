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

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkState;
import static net.sf.rubycollect4j.RubyCollections.Hash;
import static net.sf.rubycollect4j.RubyCollections.newRubyArray;
import static net.sf.rubycollect4j.RubyCollections.newRubyLazyEnumerator;
import static net.sf.rubycollect4j.RubyCollections.ra;
import static net.sf.rubycollect4j.RubyCollections.range;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import net.sf.rubycollect4j.RubyArray;
import net.sf.rubycollect4j.RubyLazyEnumerator;
import net.sf.rubycollect4j.block.TransformBlock;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.github.wnameless.nullproof.annotation.AcceptNull;
import com.github.wnameless.nullproof.annotation.RejectNull;
import com.google.common.base.Objects;
import com.google.common.collect.ArrayListMultimap;
import com.google.common.collect.Multimap;

/**
 * 
 * {@link WorkbookReader} is a wrapper to Apache POI. It tends to provide
 * friendly APIs for workbook reading.
 * 
 */
@RejectNull
public final class WorkbookReader {

  private static final Logger logger = Logger.getLogger(WorkbookReader.class
      .getName());

  private final Workbook workbook;
  private Sheet sheet;
  private final List<String> header = newRubyArray();
  private boolean hasHeader = true;
  private boolean isClosed = false;
  private FileInputStream fis;

  /**
   * Returns a {@link WorkbookReader} by given path.
   * 
   * @param path
   *          of a workbook
   * @return a {@link WorkbookReader}
   */
  public static WorkbookReader open(String path) {
    return new WorkbookReader(path);
  }

  /**
   * Returns a {@link WorkbookReader} by given file.
   * 
   * @param file
   *          of a workbook
   * @return a {@link WorkbookReader}
   */
  public static WorkbookReader open(File file) {
    return new WorkbookReader(file);
  }

  /**
   * Returns a {@link WorkbookReader} by given {@link Workbook}.
   * 
   * @param workbook
   *          a {@link Workbook}
   * @return a {@link WorkbookReader}
   */
  public static WorkbookReader open(Workbook workbook) {
    return new WorkbookReader(workbook);
  }

  /**
   * Creates a {@link WorkbookReader} by given path. Assumes there is a header
   * included in the spreadsheet.
   * 
   * @param path
   *          of a workbook
   */
  public WorkbookReader(String path) {
    workbook = createWorkbook(new File(path));
    sheet = workbook.getSheetAt(0);
    setHeader();
  }

  /**
   * Creates a {@link WorkbookReader} by given path.
   * 
   * @param path
   *          of a workbook
   * @param hasHeader
   *          true if spreadsheet gets a header, false otherwise
   * @deprecated use {@link #withoutHeader() } instead
   */
  @Deprecated
  public WorkbookReader(String path, boolean hasHeader) {
    this.hasHeader = hasHeader;
    workbook = createWorkbook(new File(path));
    sheet = workbook.getSheetAt(0);
    setHeader();
  }

  /**
   * Creates a {@link WorkbookReader} by given File. Assumes there is a header
   * included in the spreadsheet.
   * 
   * @param file
   *          of a workbook
   */
  public WorkbookReader(File file) {
    workbook = createWorkbook(file);
    sheet = workbook.getSheetAt(0);
    setHeader();
  }

  /**
   * Creates a {@link WorkbookReader} by given File.
   * 
   * @param file
   *          of a workbook
   * @param hasHeader
   *          true if spreadsheet gets a header, false otherwise
   * @deprecated use {@link #withoutHeader() } instead
   */
  @Deprecated
  public WorkbookReader(File file, boolean hasHeader) {
    this.hasHeader = hasHeader;
    workbook = createWorkbook(file);
    sheet = workbook.getSheetAt(0);
    setHeader();
  }

  /**
   * Creates a {@link WorkbookReader} by given {@link Workbook}. Assumes there
   * is a header included in the spreadsheet.
   * 
   * @param workbook
   *          a {@link Workbook}
   */
  public WorkbookReader(Workbook workbook) {
    this.workbook = workbook;
    if (workbook.getNumberOfSheets() == 0)
      workbook.createSheet();
    sheet = workbook.getSheetAt(0);
    hasHeader = true;
    setHeader();
  }

  /**
   * Creates a {@link WorkbookReader} by given Workbook.
   * 
   * @param workbook
   *          a {@link Workbook}
   * @param hasHeader
   *          true if spreadsheet gets a header, false otherwise
   * @deprecated use {@link #withoutHeader() } instead
   */
  @Deprecated
  public WorkbookReader(Workbook workbook, boolean hasHeader) {
    this.workbook = workbook;
    this.hasHeader = hasHeader;
    sheet = workbook.getSheetAt(0);
    setHeader();
  }

  private Workbook createWorkbook(File file) {
    try {
      fis = new FileInputStream(file);
      return WorkbookFactory.create(fis);
    } catch (Exception e) {
      logger.log(Level.SEVERE, null, e);
      throw new RuntimeException(e);
    }
  }

  /**
   * Mentions there is a header in this sheet.
   * 
   * @return this {@link WorkbookReader}
   */
  public WorkbookReader withHeader() {
    hasHeader = true;
    setHeader();
    return this;
  }

  /**
   * Mentions there is no header in this sheet.
   * 
   * @return this {@link WorkbookReader}
   */
  public WorkbookReader withoutHeader() {
    hasHeader = false;
    setHeader();
    return this;
  }

  private void setHeader() {
    header.clear();
    Iterator<Row> rows = sheet.rowIterator();
    if (rows.hasNext() && hasHeader)
      header.addAll(rowToRubyArray(rows.next()));
  }

  /**
   * Returns the backing POI {@link Workbook}.
   * 
   * @return the POI {@link Workbook}
   */
  public Workbook getWorkbook() {
    return workbook;
  }

  /**
   * Manually closes the Workbook file.
   */
  public void close() {
    try {
      if (fis != null)
        fis.close();
    } catch (IOException e) {
      logger.log(Level.SEVERE, null, e);
      throw new RuntimeException(e);
    }
    isClosed = true;
  }

  /**
   * Returns a List which contains all header fields.
   * 
   * @return a String List
   */
  public List<String> getHeader() {
    checkState(!isClosed, "Workbook has been closed.");
    return RubyArray.copyOf(header);
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
    checkState(!isClosed, "Workbook has been closed.");
    List<String> sheets = newRubyArray();
    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
      sheets.add(workbook.getSheetName(i));
    }
    return sheets;
  }

  /**
   * Turns this {@link WorkbookReader} to certain sheet. Sheet names can be
   * found by {@link #getAllSheetNames}.
   * 
   * @param index
   *          of a sheet
   * @return this {@link WorkbookReader}
   */
  public WorkbookReader turnToSheet(int index) {
    checkState(!isClosed, "Workbook has been closed.");
    sheet = workbook.getSheetAt(index);
    setHeader();
    return this;
  }

  /**
   * Turns this {@link WorkbookReader} to certain sheet. Sheet names can be
   * found by {@link #getAllSheetNames}.
   * 
   * @param sheetName
   *          name of a sheet
   * @return this {@link WorkbookReader}
   */
  public WorkbookReader turnToSheet(String sheetName) {
    checkArgument(getAllSheetNames().contains(sheetName),
        "Sheet name is not found.");
    return turnToSheet(getAllSheetNames().indexOf(sheetName));
  }

  /**
   * Turns this {@link WorkbookReader} to certain sheet. Sheet names can be
   * found by {@link #getAllSheetNames}.
   * 
   * @param index
   *          of a sheet
   * @param hasHeader
   *          true if spreadsheet gets a header, false otherwise
   * @return this {@link WorkbookReader}
   */
  public WorkbookReader turnToSheet(int index, boolean hasHeader) {
    checkState(!isClosed, "Workbook has been closed.");
    this.hasHeader = hasHeader;
    sheet = workbook.getSheetAt(index);
    setHeader();
    return this;
  }

  /**
   * Turns this {@link WorkbookReader} to certain sheet. Sheet names can be
   * found by {@link #getAllSheetNames}.
   * 
   * @param sheetName
   *          name of a sheet
   * @param hasHeader
   *          true if spreadsheet gets a header, false otherwise
   * @return this {@link WorkbookReader}
   */
  public WorkbookReader turnToSheet(String sheetName, boolean hasHeader) {
    checkArgument(getAllSheetNames().contains(sheetName),
        "Sheet name is not found.");
    return turnToSheet(getAllSheetNames().indexOf(sheetName), hasHeader);
  }

  /**
   * Converts the spreadsheet to CSV by a String Iterable.
   * 
   * @return a String Iterable
   */
  public Iterable<String> toCSV() {
    checkState(!isClosed, "Workbook has been closed.");
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
    checkState(!isClosed, "Workbook has been closed.");
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
    checkState(!isClosed, "Workbook has been closed.");
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
    checkState(!isClosed, "Workbook has been closed.");
    checkState(hasHeader, "Header is not provided.");
    return newRubyLazyEnumerator(sheet).map(
        new TransformBlock<Row, Map<String, String>>() {

          @SuppressWarnings("unchecked")
          @Override
          public Map<String, String> yield(Row item) {
            return Hash(ra(getHeader()).zip(rowToRubyArray(item)));
          }

        }).drop(1);
  }

  private RubyArray<String> rowToRubyArray(Row row) {
    return rowToRubyArray(row, false);
  }

  private RubyArray<String> rowToRubyArray(final Row row, boolean isCSV) {
    int colNum;
    if (hasHeader)
      colNum = sheet.rowIterator().next().getLastCellNum();
    else
      colNum = row.getLastCellNum();

    return range(0, colNum - 1).map(new TransformBlock<Integer, Cell>() {

      @Override
      public Cell yield(Integer item) {
        return row.getCell(item);
      }

    }).map(cell2Str(isCSV));
  }

  private TransformBlock<Cell, String> cell2Str(final boolean isCSV) {
    return new TransformBlock<Cell, String>() {

      @AcceptNull
      @Override
      public String yield(Cell item) {
        if (item == null)
          return "";

        item.setCellType(CELL_TYPE_STRING);
        String val = item.toString();
        if (isCSV && val.contains(",")) {
          val = val.replaceAll("\"", "\"\"");
          return '"' + val + '"';
        }
        return val;
      }

    };
  }

  /**
   * Converts this {@link WorkbookReader} to a {@link WorkbookWriter}.
   * 
   * @return a {@link WorkbookWriter}
   */
  public WorkbookWriter toWriter() {
    return new WorkbookWriter(workbook);
  }

  /**
   * Returns a {@link Multimap} which represents the content of this workbook.
   * Each sheet name is used as the key, and the value is a Collection of String
   * List which contains all fields of a row.
   * 
   * @return a {@link Multimap}
   */
  public Multimap<String, List<String>> toMultimap() {
    Multimap<String, List<String>> content = ArrayListMultimap.create();

    String currentSheet = getCurrentSheetName();
    boolean currentHeader = hasHeader;

    for (String sheetName : getAllSheetNames()) {
      turnToSheet(sheetName);
      withoutHeader();
      for (List<String> row : toLists()) {
        content.put(sheetName, row);
      }
    }

    turnToSheet(currentSheet);
    hasHeader = currentHeader;

    return content;
  }

  @Override
  public boolean equals(Object o) {
    if (o instanceof WorkbookReader) {
      WorkbookReader reader = (WorkbookReader) o;
      return Objects.equal(toMultimap(), reader.toMultimap());
    }
    return false;
  }

  @Override
  public int hashCode() {
    return Objects.hashCode(toMultimap());
  }

  @Override
  public String toString() {
    return Objects.toStringHelper(getClass()).addValue(toMultimap()).toString();
  }

}
