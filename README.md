[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.wnameless/workbook-accessor/badge.svg)](https://maven-badges.herokuapp.com/maven-central/com.github.wnameless/workbook-accessor)

workbook-accessor
=============
A friendly Java workbook writer and reader based on Apache POI

## Purpose
Sometimes, you only need to do simple jobs with workbook files(Excel, Spreadsheet...).<br>
The workbook-accessor provides you an easy and convenient way to manipulate workbooks in a breeze.

## Maven Repo
```xml
<dependency>
    <groupId>com.github.wnameless</groupId>
    <artifactId>workbook-accessor</artifactId>
    <version>1.5.0</version>
</dependency>
```
Since 1.4.0, Java 8 required.

## Quick Start
#### WorkbookReader
```java
WorkbookReader reader = WorkbookReader.open("path_to_workbook");
for(List<String> row : reader.toLists()) {
  // Print all the rows in workbook
  System.out.println(row);
}
```

#### WorkbookWriter
```java
WorkbookWriter writer = WorkbookWriter.openXLSX().setSheetName("name_of_the_sheet");
writer.addRow(123, "abc", new Date())
      .addRow(1.1, 2.2f, 33L)
      .save("path_of_the_output_file");
```

## Features
Works on multiple sheets.
```java
reader.turnToSheet("Sheet0");
writer.createAndTurnToSheet("NewSheet");
```

More than one way to iterate over the rows of a sheet.
```java
reader.toCSV();    // Returns a Iterable<String>
reader.toLists();  // Returns a Iterable<List<String>>
reader.toArrays(); // Returns a Iterable<String[]>
reader.toMaps();   // Returns a Iterable<Map<String, String>>
```

Add a new row to a sheet in different ways.
```java
writer.addRow(Arrays.asList("a", "b", "c")); // Accepts any Iterable
writer.addRow(123, "abc", new Date());       // Object VarArgs
```

Retrieve whole content at once.
```java
WorkbookReader.open("path_to_workbook").toMultimap(); // {Sheet1=>[[123, abc, !@#], [4, 5, 6], [d, e, f]], Sheet2=>[...], ...}
```

Interchangeable reader and writer.

```java
reader.toWriter();
writer.toReader();
```
Get the backing Workbook or turn the Workbook into a byte array easily.
```java
Workbook workbook;
workbook = reader.getWorkbook();
workbook = writer.getWorkbook();

byte[] workbookBytes;
workbookBytes = reader.toBytes();
workbookBytes = writer.toBytes();
```