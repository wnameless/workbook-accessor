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

import static net.sf.rubycollect4j.RubyCollections.newRubyArray;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * WorkbookWriter is a wrapper to Apache POI. It tends to provide friendly APIs
 * for workbook writing.
 * 
 */
public final class WorkbookWriter {

  private static final Logger logger = Logger.getLogger(WorkbookWriter.class
      .getName());

  private final Workbook wb;
  private Sheet sheet;

  /**
   * Creates a WorkbookWriter. Default sheet name is Sheet0 and xls file is
   * used.
   */
  public WorkbookWriter() {
    wb = new HSSFWorkbook();
    sheet = wb.createSheet();
  }

  /**
   * Creates a WorkbookWriter by given Workbook.
   * 
   * @param wb
   *          a Workbook
   */
  public WorkbookWriter(Workbook wb) {
    if (wb == null)
      throw new NullPointerException("Workbook can't be null.");

    this.wb = wb;
    if (wb.getNumberOfSheets() == 0)
      wb.createSheet();
    sheet = wb.getSheetAt(0);
  }

  /**
   * Creates a WorkbookWriter and creates a new sheet by given name.
   * 
   * @param sheetName
   *          name of the sheet
   */
  public WorkbookWriter(String sheetName) {
    if (sheetName == null)
      throw new NullPointerException("Sheet name can't be null.");

    wb = new HSSFWorkbook();
    sheet = wb.createSheet(sheetName);
  }

  /**
   * Creates a WorkbookWriter.
   * 
   * @param xlsx
   *          true if an xlsx file is used, otherwise false
   */
  public WorkbookWriter(boolean xlsx) {
    if (xlsx)
      wb = new HSSFWorkbook();
    else
      wb = new XSSFWorkbook();

    sheet = wb.createSheet();
  }

  /**
   * Creates a WorkbookWriter and creates a new sheet by given name.
   * 
   * @param sheetName
   *          name of the sheet
   * @param xlsx
   *          true if an xlsx file is used, otherwise false
   */
  public WorkbookWriter(String sheetName, boolean xlsx) {
    if (sheetName == null)
      throw new NullPointerException("Sheet name can't be null.");

    if (xlsx)
      wb = new HSSFWorkbook();
    else
      wb = new XSSFWorkbook();

    sheet = wb.createSheet(sheetName);
  }

  /**
   * Returns the backing POI Workbook.
   * 
   * @return the POI Workbook
   */
  public Workbook getWorkbook() {
    return wb;
  }

  /**
   * Returns the name of current sheet.
   * 
   * @return the name of the sheet
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
    List<String> sheets = newRubyArray();
    for (int i = 0; i < wb.getNumberOfSheets(); i++) {
      sheets.add(wb.getSheetName(i));
    }
    return sheets;
  }

  /**
   * Creates a new sheet.
   * 
   * @param sheetName
   *          name of a sheet
   * @return this WorkbookWriter
   */
  public WorkbookWriter createSheet(String sheetName) {
    if (sheetName == null)
      throw new NullPointerException("Sheet name can't be null.");

    wb.createSheet(sheetName);
    return this;
  }

  /**
   * Turns this WorkbookWriter to certain sheet. Sheets can be found by
   * getAllSheetNames().
   * 
   * @param index
   *          of a sheet
   * @return this WorkbookWriter
   */
  public WorkbookWriter turnToSheet(int index) {
    sheet = wb.getSheetAt(index);
    return this;
  }

  /**
   * Creates a new sheet and turns this WorkbookWriter to the sheet.
   * 
   * @param name
   *          of a sheet
   * @return this WorkbookWriter
   */
  public WorkbookWriter turnToSheet(String name) {
    if (name == null)
      throw new NullPointerException("Name can't be null.");

    return turnToSheet(getAllSheetNames().indexOf(name));
  }

  /**
   * Turns this WorkbookWriter to certain sheet. Sheets can be found by
   * getAllSheetNames().
   * 
   * @param name
   *          of a sheet
   * @return this WorkbookWriter
   */
  public WorkbookWriter createAndTurnToSheet(String name) {
    if (name == null)
      throw new NullPointerException("Name can't be null.");

    sheet = wb.createSheet(name);
    return this;
  }

  /**
   * Adds a row to the sheet.
   * 
   * @param fields
   *          an Iterable of Object
   * @return this WorkbookWriter
   */
  public WorkbookWriter addRow(Iterable<? extends Object> fields) {
    if (fields == null)
      throw new NullPointerException("Fields can't be null.");

    Row row;
    if (sheet.getLastRowNum() == 0 && sheet.getPhysicalNumberOfRows() == 0)
      row = sheet.createRow(0);
    else
      row = sheet.createRow(sheet.getLastRowNum() + 1);

    int i = 0;
    for (Object o : fields) {
      Cell cell = row.createCell(i);
      if (o != null) {
        if (o instanceof Boolean) {
          cell.setCellValue((Boolean) o);
        } else if (o instanceof Calendar) {
          cell.setCellValue((Calendar) o);
        } else if (o instanceof Date) {
          cell.setCellValue((Date) o);
        } else if (o instanceof Double) {
          cell.setCellValue((Double) o);
        } else if (o instanceof RichTextString) {
          cell.setCellValue((RichTextString) o);
        } else if (o instanceof Hyperlink) {
          cell.setHyperlink((Hyperlink) o);
        } else {
          cell.setCellValue(o.toString());
        }
      }
      i++;
    }
    return this;
  }

  /**
   * Adds a row to the sheet.
   * 
   * @param fields
   *          an array of Object
   * @return this WorkbookWriter
   */
  public WorkbookWriter addRow(Object... fields) {
    if (fields == null)
      throw new NullPointerException("Fields can't be null.");

    return addRow(Arrays.asList(fields));
  }

  /**
   * Saves this WorkbookWriter to a file.
   * 
   * @param path
   *          of the output file
   * @return a saved File
   */
  public File save(String path) {
    if (path == null)
      throw new NullPointerException("Path can't be null.");

    try {
      FileOutputStream out = new FileOutputStream(path);
      wb.write(out);
      out.close();
    } catch (IOException e) {
      logger.log(Level.SEVERE, null, e);
      throw new RuntimeException(e);
    }
    return new File(path);
  }

}
