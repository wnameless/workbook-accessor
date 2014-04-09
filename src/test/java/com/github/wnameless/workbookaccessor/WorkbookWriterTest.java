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

import static net.sf.rubycollect4j.RubyCollections.ra;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import net.sf.rubycollect4j.RubyFile;

import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import com.google.common.testing.NullPointerTester;

public class WorkbookWriterTest {

  private static final String BASE_DIR = "src/test/resources";
  private WorkbookWriter writer;

  @Before
  public void setUp() throws Exception {
    writer = new WorkbookWriter();
  }

  @Test
  public void testAllConstructorsNPE() {
    new NullPointerTester().testAllPublicConstructors(WorkbookWriter.class);
  }

  @Test
  public void testAllPublicMethodsNPE() {
    new NullPointerTester().testAllPublicInstanceMethods(writer);
  }

  @Test
  public void testConstructor() {
    assertTrue(writer instanceof WorkbookWriter);
    assertTrue(new WorkbookWriter("test") instanceof WorkbookWriter);
    assertTrue(new WorkbookWriter(true) instanceof WorkbookWriter);
    assertTrue(new WorkbookWriter(false) instanceof WorkbookWriter);
    assertTrue(new WorkbookWriter("test", true) instanceof WorkbookWriter);
    assertTrue(new WorkbookWriter("test", false) instanceof WorkbookWriter);
    assertTrue(new WorkbookWriter(new HSSFWorkbook()) instanceof WorkbookWriter);
    Workbook wb = new HSSFWorkbook();
    wb.createSheet();
    assertTrue(new WorkbookWriter(wb) instanceof WorkbookWriter);
  }

  @Test
  public void testGetWorkbook() {
    assertTrue(writer.getWorkbook() instanceof Workbook);
  }

  @Test
  public void testGetCurrentSheetName() {
    assertEquals("Sheet0", writer.getCurrentSheetName());
  }

  @Test
  public void testGetAllSheetNames() {
    assertEquals(ra("Sheet0"), writer.getAllSheetNames());
  }

  @Test
  public void testCreateSheet() {
    writer.createSheet("test");
    assertTrue(writer.getAllSheetNames().contains("test"));
    assertEquals("Sheet0", writer.getCurrentSheetName());
  }

  @Test(expected = IllegalArgumentException.class)
  public void testCreateSheetException() {
    writer.createSheet("test");
    writer.createSheet("test");
  }

  @Test
  public void testTurnToSheet() {
    writer.createSheet("test");
    writer.turnToSheet("test");
    assertEquals("test", writer.getCurrentSheetName());
    writer.turnToSheet(0);
    assertEquals("Sheet0", writer.getCurrentSheetName());
  }

  @Test(expected = IllegalArgumentException.class)
  public void testTurnToSheetException() {
    writer.turnToSheet("hahaha");
  }

  @Test
  public void testCreateAndTurnToSheet() {
    writer.createAndTurnToSheet("test");
    assertEquals("test", writer.getCurrentSheetName());
  }

  @Test(expected = IllegalArgumentException.class)
  public void testCreateAndTurnToSheetException() {
    writer.createAndTurnToSheet("test");
    writer.createAndTurnToSheet("test");
  }

  @Test
  public void testAddRow() {
    Calendar cal = Calendar.getInstance();
    Date date = new Date();
    writer.addRow("def");
    writer.addRow(
        null,
        true,
        cal,
        date,
        1.1,
        new HSSFRichTextString("Hello!"),
        new HSSFWorkbook().getCreationHelper().createHyperlink(
            Hyperlink.LINK_URL), "abc");
    assertEquals("def", writer.getWorkbook().getSheetAt(0).rowIterator().next()
        .cellIterator().next().getStringCellValue());
  }

  @Test
  public void testSave() throws InvalidFormatException, IOException {
    writer.addRow("abc", "def");
    writer.save(RubyFile.join(BASE_DIR, "test.xls"));
    WorkbookReader reader =
        new WorkbookReader(RubyFile.join(BASE_DIR, "test.xls"), false);
    assertEquals("abc,def", reader.toCSV().iterator().next());
    reader.close();
    RubyFile.delete(RubyFile.join(BASE_DIR, "test.xls"));
  }

}
