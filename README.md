[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.wnameless/workbook-accessor/badge.svg)](https://maven-badges.herokuapp.com/maven-central/com.github.wnameless/workbook-accessor)

workbook-accessor
=============
A friendly workbook writer and reader for Java based on Apache POI

## Purpose
Sometimes, you only need to do simple jobs with workbook files(Excel, Spreadsheet...).<br>
The workbook-accessor provides you an easy and convenient way to manipulate workbooks in a breeze.

## Maven Repo
```xml
<dependency>
    <groupId>com.github.wnameless</groupId>
    <artifactId>workbook-accessor</artifactId>
    <version>1.3.3</version>
</dependency>
```

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

More than one way to interate over the rows of a sheet.
```java
reader.toCSV();    // Returns a Iterable<String>
reader.toLists();  // Returns a Iterable<List<String>>
reader.toArrays(); // Returns a Iterable<String[]>
reader.toMaps();   // Returns a Iterable<Map<String, String>>
```

More than one way to add a new row to a sheet.
```java

writer.addRow(Arrays.asList("a", "b", "c")); // Accepts any Iterable
writer.addRow(123, "abc", new Date());       // Object VarArgs
```

Retrieve whole content at once.
```java
WorkbookReader.open("path_to_workbook").toMultimap(); // {Sheet1=>[[123, abc, !@#], [4, 5, 6], [d, e, f]], Sheet2=>[...], ...}
```

Interchangeable reader and writer
```java
reader.toWriter();
writer.toReader();
```
