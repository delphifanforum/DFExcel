# DFExcel

A powerful Delphi unit for seamless interaction with Microsoft Excel, providing a comprehensive set of procedures and functions to manipulate Excel files programmatically. Compatible with modern Delphi versions.

## Overview

DFExcel (Delphi-Fast Excel) simplifies Excel automation in Delphi applications by providing a clean, object-oriented interface to Microsoft Excel's COM objects. It handles common operations like reading/writing data, formatting cells, managing worksheets, and performing calculations, while abstracting away the complexities of the Excel object model.

## Features

- **Simple Excel Instance Management**: Create, open, save, and close Excel workbooks with minimal code
- **Comprehensive Cell Operations**: Read and write values of any type (string, integer, float, boolean, date)
- **Worksheet Management**: Create, delete, rename, and navigate between worksheets
- **Range Manipulation**: Select, copy, paste, and format cell ranges
- **Formula Support**: Apply formulas and perform calculations
- **Formatting Options**: Apply number formats, fonts, borders, colors
- **Data Import/Export**: Import from and export to various formats including CSV and text files
- **Chart Creation**: Generate charts from data ranges
- **Error Handling**: Robust error capture and reporting system
- **Modern Delphi Support**: Compatible with current Delphi versions including Delphi 10.x, 11.x and beyond

## Requirements

- Delphi 10.x or higher (also works with older versions with minor modifications)
- Microsoft Excel installed on the system
- Windows operating system

## Installation

1. Clone this repository or download the source code
2. Add the `DFExcel.pas` unit to your Delphi project
3. Add `DFExcel` to your uses clause
4. Make sure your project includes the required COM dependencies

## Quick Start

```delphi
uses
  DFExcel;

var
  Excel: TDFExcelApplication;
  Workbook: TDFExcelWorkbook;
  Sheet: TDFExcelWorksheet;
begin
  // Initialize Excel application
  Excel := TDFExcelApplication.Create;
  try
    // Create a new workbook
    Workbook := Excel.CreateWorkbook;
    
    // Get the first worksheet
    Sheet := Workbook.Worksheets[1];
    
    // Write data to cells
    Sheet.Cells[1, 1].Value := 'Product';
    Sheet.Cells[1, 2].Value := 'Quantity';
    Sheet.Cells[1, 3].Value := 'Price';
    
    Sheet.Cells[2, 1].Value := 'Widget';
    Sheet.Cells[2, 2].Value := 100;
    Sheet.Cells[2, 3].Value := 19.99;
    
    // Apply formula
    Sheet.Cells[2, 4].Formula := '=B2*C2';
    
    // Format header row
    Sheet.Range['A1:D1'].Font.Bold := True;
    
    // Save workbook
    Workbook.SaveAs('SampleReport.xlsx');
  finally
    // Clean up resources
    Excel.Free;
  end;
end;
```

## API Reference

### TDFExcelApplication

The main class representing an Excel application instance.

#### Methods

- `Create(Visible: Boolean = False)`: Creates a new Excel application instance
- `CreateWorkbook`: Creates a new workbook
- `OpenWorkbook(FileName: string)`: Opens an existing workbook
- `Quit`: Closes Excel application
- `DisplayAlerts(Value: Boolean)`: Controls whether Excel displays alerts
- `ScreenUpdating(Value: Boolean)`: Controls screen updating for performance
- `Calculate`: Calculates all open workbooks
- `Evaluate(Name: string)`: Evaluates a named formula or range

#### Properties

- `Visible`: Sets Excel visibility (boolean property)
- `ActiveWorkbook`: Gets the currently active workbook
- `Workbooks[Index: Integer]`: Access workbooks by index
- `WorkbookCount`: Returns the number of open workbooks

### TDFExcelWorkbook

Represents an Excel workbook.

#### Methods

- `AddWorksheet(Name: string = '')`: Adds a new worksheet
- `WorksheetByName(Name: string)`: Access worksheet by name
- `Save`: Saves the workbook
- `SaveAs(FileName: string; FileFormat: Integer = xlOpenXMLWorkbook)`: Saves the workbook with the specified name
- `Close(SaveChanges: Boolean = True)`: Closes the workbook

#### Properties

- `Worksheets[Index: Integer]`: Access worksheets by index
- `WorksheetCount`: Returns the number of worksheets
- `FileName`: Returns the full path of the workbook

### TDFExcelWorksheet

Represents an Excel worksheet.

#### Methods

- `Cells[Row, Col: Integer]`: Access individual cells
- `Range[Address: string]`: Access a range of cells
- `CopySheet(After: OleVariant)`: Copy this worksheet
- `Delete`: Delete this worksheet
- `Activate`: Make this worksheet active
- `InsertRow(RowIndex: Integer)`: Insert a row
- `InsertColumn(ColIndex: Integer)`: Insert a column
- `DeleteRow(RowIndex: Integer)`: Delete a row
- `DeleteColumn(ColIndex: Integer)`: Delete a column
- `CreateChart(ChartType: TDFExcelChartType; DataRange: string)`: Create a chart
- `PrintOut(From: Integer = 0; To_: Integer = 0; Copies: Integer = 1)`: Print the worksheet
- `PrintPreview`: Preview the worksheet for printing

#### Properties

- `Name`: The worksheet name
- `UsedRange`: The range of cells that contains data
- `Visible`: Controls worksheet visibility

### TDFExcelRange

Represents a range of Excel cells.

#### Methods

- `Clear`: Clears all content and formatting
- `ClearContents`: Clears only the content
- `ClearFormats`: Clears only the formatting
- `AutoFit`: Auto-fits columns to content
- `Merge`: Merges cells in the range
- `Unmerge`: Unmerges previously merged cells
- `Copy`: Copies the range to the clipboard
- `Cut`: Cuts the range to the clipboard
- `Paste`: Pastes from the clipboard
- `PasteSpecial(Format: Integer = xlPasteAll)`: Pastes with specific options
- `Sort(Key1Range, Key2Range, ...)`: Sorts the range
- `CreateFormatCondition(...)`: Creates conditional formatting

#### Properties

- `Value`: Gets or sets the cell values (variant)
- `Formula`: Gets or sets formulas
- `HorizontalAlignment`: Gets or sets horizontal alignment
- `VerticalAlignment`: Gets or sets vertical alignment
- `Font`: Access font properties
- `Interior`: Access cell background properties
- `Borders`: Access cell border properties

### TDFExcelCell

Represents an individual Excel cell.

#### Properties

- `Value`: Gets or sets the cell value (variant)
- `Formula`: Gets or sets a formula
- `NumberFormat`: Gets or sets number format string
- `HorizontalAlignment`: Gets or sets horizontal alignment
- `VerticalAlignment`: Gets or sets vertical alignment
- `Font`: Access font properties
- `Interior`: Access cell background properties
- `Borders`: Access cell border properties

### TDFExcelChart

Represents an Excel chart.

#### Methods

- `SetSourceData(Range: OleVariant; PlotBy: Integer = xlColumns)`: Set the chart data source
- `Export(FileName: string; FilterName: string = 'PNG')`: Export chart as image

#### Properties

- `ChartType`: Gets or sets the chart type
- `HasTitle`: Gets or sets whether the chart has a title
- `ChartTitle`: Access the chart title properties
- `HasLegend`: Gets or sets whether the chart has a legend

## Advanced Features

### Range Formatting

```delphi
// Apply formatting to a range
var Range: TDFExcelRange;
begin
  Range := Sheet.Range['A1:D10'];
  
  // Apply borders
  Range.Borders.LineStyle := xlContinuous;
  Range.Borders.Weight := xlThin;
  
  // Apply background color
  Range.Interior.Color := clLightBlue;
  
  // Apply text alignment
  Range.HorizontalAlignment := xlCenter;
end;
```

### Creating Charts

```delphi
// Create a chart from data
var Chart: TDFExcelChart;
begin
  // Specify data range and chart type
  Chart := Sheet.CreateChart(ctBar, 'A1:B10');
  
  // Configure chart properties
  Chart.HasTitle := True;
  Chart.ChartTitle.Text := 'Sales Report';
  Chart.HasLegend := True;
end;
```

### Batch Operations

```delphi
// Import data from array
var
  DataArray: Variant;
  i, j: Integer;
begin
  // Create a 2D array
  DataArray := VarArrayCreate([1, 100, 1, 5], varVariant);
  
  // Fill with data
  for i := 1 to 100 do
    for j := 1 to 5 do
      DataArray[i, j] := i * j;
  
  // Write entire array to Excel in one operation
  Sheet.Range['A1:E100'].Value := DataArray;
end;
```

## Performance Tips

1. **Disable screen updating** during batch operations:
```delphi
Excel.ScreenUpdating(False);
try
  // Perform operations
finally
  Excel.ScreenUpdating(True);
end;
```

2. **Use arrays for bulk operations** instead of setting cells individually:
```delphi
// Instead of:
for i := 1 to 1000 do
  Sheet.Cells[i, 1].Value := i;

// Use:
var Data: Variant;
Data := VarArrayCreate([1, 1000, 1, 1], varVariant);
for i := 1 to 1000 do
  Data[i, 1] := i;
Sheet.Range['A1:A1000'].Value := Data;
```

3. **Avoid excessive property access** across COM boundaries:
```delphi
// Instead of:
for i := 1 to 100 do begin
  Sheet.Cells[i, 1].Font.Bold := True;
  Sheet.Cells[i, 1].Font.Size := 12;
end;

// Use:
Sheet.Range['A1:A100'].Font.Bold := True;
Sheet.Range['A1:A100'].Font.Size := 12;
```

## Error Handling

DFExcel includes comprehensive error handling to gracefully manage Excel-related exceptions:

```delphi
try
  // Excel operations
  Sheet.Cells[1, 1].Value := 'Test';
except
  on E: EDFExcelException do
    ShowMessage('Excel error: ' + E.Message);
  on E: Exception do
    ShowMessage('General error: ' + E.Message);
end;
```

## Best Practices

1. Always properly clean up Excel instances to prevent memory leaks
2. Use batch operations when handling large data sets
3. Use structured references when working with Excel tables
4. Handle COM errors appropriately
5. Consider creating a separate thread for time-consuming Excel operations
6. Be aware that some operations may trigger Excel's modal dialogs

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgements

- Thanks to the Delphi community for feedback and suggestions
- Special thanks to Microsoft for documenting the Excel Object Model
