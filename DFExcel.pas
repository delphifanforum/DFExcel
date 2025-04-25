unit DFExcel;

interface

uses
  Winapi.Windows, System.SysUtils, System.Variants, System.Classes, 
  System.Win.ComObj, Winapi.ActiveX, Vcl.Dialogs, Vcl.Graphics;

type
  // Forward declarations
  TDFExcelApplication = class;
  TDFExcelWorkbook = class;
  TDFExcelWorksheet = class;
  TDFExcelRange = class;
  TDFExcelCell = class;
  TDFExcelChart = class;
  
  // Custom exceptions
  EDFExcelException = class(Exception);

  // Excel chart types
  TDFExcelChartType = (
    ctArea, ctBar, ctColumn, ctLine, ctPie, ctRadar, ctXY, ct3DArea, 
    ct3DBar, ct3DColumn, ct3DLine, ct3DPie, ct3DSurface
  );

  // Font properties type
  TDFExcelFont = class
  private
    FRange: OleVariant;
    function GetBold: Boolean;
    procedure SetBold(const Value: Boolean);
    function GetItalic: Boolean;
    procedure SetItalic(const Value: Boolean);
    function GetSize: Integer;
    procedure SetSize(const Value: Integer);
    function GetColor: TColor;
    procedure SetColor(const Value: TColor);
    function GetName: string;
    procedure SetName(const Value: string);
    function GetUnderline: Boolean;
    procedure SetUnderline(const Value: Boolean);
  public
    constructor Create(Range: OleVariant);
    property Bold: Boolean read GetBold write SetBold;
    property Italic: Boolean read GetItalic write SetItalic;
    property Size: Integer read GetSize write SetSize;
    property Color: TColor read GetColor write SetColor;
    property Name: string read GetName write SetName;
    property Underline: Boolean read GetUnderline write SetUnderline;
  end;

  // Interior (background) properties
  TDFExcelInterior = class
  private
    FRange: OleVariant;
    function GetColor: TColor;
    procedure SetColor(const Value: TColor);
    function GetPattern: Integer;
    procedure SetPattern(const Value: Integer);
  public
    constructor Create(Range: OleVariant);
    property Color: TColor read GetColor write SetColor;
    property Pattern: Integer read GetPattern write SetPattern;
  end;

  // Border properties
  TDFExcelBorders = class
  private
    FRange: OleVariant;
    function GetLineStyle: Integer;
    procedure SetLineStyle(const Value: Integer);
    function GetWeight: Integer;
    procedure SetWeight(const Value: Integer);
    function GetColor: TColor;
    procedure SetColor(const Value: TColor);
  public
    constructor Create(Range: OleVariant);
    property LineStyle: Integer read GetLineStyle write SetLineStyle;
    property Weight: Integer read GetWeight write SetWeight;
    property Color: TColor read GetColor write SetColor;
  end;

  // Cell class for individual cell operations
  TDFExcelCell = class
  private
    FCell: OleVariant;
    FFont: TDFExcelFont;
    FInterior: TDFExcelInterior;
    FBorders: TDFExcelBorders;
    function GetValue: Variant;
    procedure SetValue(const Value: Variant);
    function GetFormula: string;
    procedure SetFormula(const Value: string);
    function GetNumberFormat: string;
    procedure SetNumberFormat(const Value: string);
    function GetHorizontalAlignment: Integer;
    procedure SetHorizontalAlignment(const Value: Integer);
    function GetVerticalAlignment: Integer;
    procedure SetVerticalAlignment(const Value: Integer);
  public
    constructor Create(Cell: OleVariant);
    destructor Destroy; override;
    property Value: Variant read GetValue write SetValue;
    property Formula: string read GetFormula write SetFormula;
    property NumberFormat: string read GetNumberFormat write SetNumberFormat;
    property HorizontalAlignment: Integer read GetHorizontalAlignment write SetHorizontalAlignment;
    property VerticalAlignment: Integer read GetVerticalAlignment write SetVerticalAlignment;
    property Font: TDFExcelFont read FFont;
    property Interior: TDFExcelInterior read FInterior;
    property Borders: TDFExcelBorders read FBorders;
  end;

  // Range class for operations on cell ranges
  TDFExcelRange = class
  private
    FRange: OleVariant;
    FFont: TDFExcelFont;
    FInterior: TDFExcelInterior;
    FBorders: TDFExcelBorders;
    function GetValue: Variant;
    procedure SetValue(const Value: Variant);
    function GetFormula: string;
    procedure SetFormula(const Value: string);
    function GetHorizontalAlignment: Integer;
    procedure SetHorizontalAlignment(const Value: Integer);
    function GetVerticalAlignment: Integer;
    procedure SetVerticalAlignment(const Value: Integer);
  public
    constructor Create(Range: OleVariant);
    destructor Destroy; override;
    procedure Clear;
    procedure ClearContents;
    procedure ClearFormats;
    procedure AutoFit;
    procedure Merge;
    procedure Unmerge;
    procedure Copy;
    procedure Cut;
    procedure Paste;
    procedure PasteSpecial(Format: Integer = -4163); // xlPasteAll = -4163
    procedure Sort(Key1Range: string; Order1: Integer = 1; Key2Range: string = ''; 
                  Order2: Integer = 1; Header: Integer = 2);
    function CreateFormatCondition(ConditionType: Integer; Operator: Integer; 
                                  Formula1: string; Formula2: string = ''): OleVariant;
    property Value: Variant read GetValue write SetValue;
    property Formula: string read GetFormula write SetFormula;
    property HorizontalAlignment: Integer read GetHorizontalAlignment write SetHorizontalAlignment;
    property VerticalAlignment: Integer read GetVerticalAlignment write SetVerticalAlignment;
    property Font: TDFExcelFont read FFont;
    property Interior: TDFExcelInterior read FInterior;
    property Borders: TDFExcelBorders read FBorders;
  end;

  // Chart class for Excel charts
  TDFExcelChart = class
  private
    FChart: OleVariant;
    function GetChartType: TDFExcelChartType;
    procedure SetChartType(const Value: TDFExcelChartType);
    function GetHasTitle: Boolean;
    procedure SetHasTitle(const Value: Boolean);
    function GetChartTitle: OleVariant;
    function GetHasLegend: Boolean;
    procedure SetHasLegend(const Value: Boolean);
  public
    constructor Create(Chart: OleVariant);
    procedure Export(FileName: string; FilterName: string = 'PNG');
    procedure SetSourceData(Range: OleVariant; PlotBy: Integer = 2); // xlColumns = 2
    property ChartType: TDFExcelChartType read GetChartType write SetChartType;
    property HasTitle: Boolean read GetHasTitle write SetHasTitle;
    property ChartTitle: OleVariant read GetChartTitle;
    property HasLegend: Boolean read GetHasLegend write SetHasLegend;
  end;

  // Worksheet class for Excel worksheets
  TDFExcelWorksheet = class
  private
    FWorksheet: OleVariant;
    function GetName: string;
    procedure SetName(const Value: string);
    function GetCell(Row, Col: Integer): TDFExcelCell;
    function GetRange(const Address: string): TDFExcelRange;
    function GetUsedRange: TDFExcelRange;
    function GetVisible: Boolean;
    procedure SetVisible(const Value: Boolean);
  public
    constructor Create(Worksheet: OleVariant);
    destructor Destroy; override;
    function CopySheet(After: OleVariant): TDFExcelWorksheet;
    procedure Delete;
    procedure Activate;
    procedure InsertRow(RowIndex: Integer);
    procedure InsertColumn(ColIndex: Integer);
    procedure DeleteRow(RowIndex: Integer);
    procedure DeleteColumn(ColIndex: Integer);
    function CreateChart(ChartType: TDFExcelChartType; DataRange: string): TDFExcelChart;
    procedure PrintOut(From: Integer = 0; To_: Integer = 0; Copies: Integer = 1);
    procedure PrintPreview;
    property Name: string read GetName write SetName;
    property Cells[Row, Col: Integer]: TDFExcelCell read GetCell; default;
    property Range[const Address: string]: TDFExcelRange read GetRange;
    property UsedRange: TDFExcelRange read GetUsedRange;
    property Visible: Boolean read GetVisible write SetVisible;
  end;

  // Workbook class for Excel workbooks
  TDFExcelWorkbook = class
  private
    FWorkbook: OleVariant;
    FSheets: TList;
    function GetWorksheet(Index: Integer): TDFExcelWorksheet;
    function GetWorksheetCount: Integer;
    function GetFileName: string;
  public
    constructor Create(Workbook: OleVariant);
    destructor Destroy; override;
    function AddWorksheet(const Name: string = ''): TDFExcelWorksheet;
    function WorksheetByName(const Name: string): TDFExcelWorksheet;
    procedure Save;
    procedure SaveAs(const FileName: string; FileFormat: Integer = 51); // xlOpenXMLWorkbook = 51
    procedure Close(SaveChanges: Boolean = True);
    property Worksheets[Index: Integer]: TDFExcelWorksheet read GetWorksheet;
    property WorksheetCount: Integer read GetWorksheetCount;
    property FileName: string read GetFileName;
  end;

  // Main Excel application class
  TDFExcelApplication = class
  private
    FExcel: OleVariant;
    FWorkbooks: TList;
    function GetVisible: Boolean;
    procedure SetVisible(const Value: Boolean);
    function GetActiveWorkbook: TDFExcelWorkbook;
    function GetWorkbook(Index: Integer): TDFExcelWorkbook;
    function GetWorkbookCount: Integer;
  public
    constructor Create(Visible: Boolean = False);
    destructor Destroy; override;
    function CreateWorkbook: TDFExcelWorkbook;
    function OpenWorkbook(const FileName: string): TDFExcelWorkbook;
    procedure Quit;
    procedure DisplayAlerts(Value: Boolean);
    procedure ScreenUpdating(Value: Boolean);
    procedure Calculate;
    function Evaluate(const Name: string): Variant;
    property Visible: Boolean read GetVisible write SetVisible;
    property ActiveWorkbook: TDFExcelWorkbook read GetActiveWorkbook;
    property Workbooks[Index: Integer]: TDFExcelWorkbook read GetWorkbook;
    property WorkbookCount: Integer read GetWorkbookCount;
  end;

// Excel constants
const
  // Excel constants - Horizontal Alignment
  xlHAlignCenter = -4108;
  xlHAlignCenterAcrossSelection = 7;
  xlHAlignDistributed = -4117;
  xlHAlignFill = 5;
  xlHAlignGeneral = 1;
  xlHAlignJustify = -4130;
  xlHAlignLeft = -4131;
  xlHAlignRight = -4152;
  
  // Excel constants - Vertical Alignment
  xlVAlignBottom = -4107;
  xlVAlignCenter = -4108;
  xlVAlignDistributed = -4117;
  xlVAlignJustify = -4130;
  xlVAlignTop = -4160;
  
  // Excel constants - Line Style
  xlContinuous = 1;
  xlDash = -4115;
  xlDashDot = 4;
  xlDashDotDot = 5;
  xlDot = -4118;
  xlDouble = -4119;
  xlLineStyleNone = -4142;
  xlSlantDashDot = 13;
  
  // Excel constants - Weight
  xlHairline = 1;
  xlMedium = -4138;
  xlThick = 4;
  xlThin = 2;
  
  // Excel constants - Pattern
  xlPatternAutomatic = -4105;
  xlPatternChecker = 9;
  xlPatternCrissCross = 16;
  xlPatternDown = -4121;
  xlPatternGray16 = 17;
  xlPatternGray25 = -4124;
  xlPatternGray50 = -4125;
  xlPatternGray75 = -4126;
  xlPatternGray8 = 18;
  xlPatternGrid = 15;
  xlPatternHorizontal = -4128;
  xlPatternLightDown = 13;
  xlPatternLightHorizontal = 11;
  xlPatternLightUp = 14;
  xlPatternLightVertical = 12;
  xlPatternNone = -4142;
  xlPatternSemiGray75 = 10;
  xlPatternSolid = 1;
  xlPatternUp = -4162;
  xlPatternVertical = -4166;
  
  // Excel constants - File Format
  xlExcel12 = 50;               // Excel 2010
  xlOpenXMLWorkbook = 51;       // Excel 2007+ (.xlsx)
  xlOpenXMLWorkbookMacroEnabled = 52; // Excel 2007+ macro-enabled (.xlsm)
  xlExcel8 = 56;                // Excel 97-2003 (.xls)
  xlCSV = 6;                    // CSV format
  xlTextWindows = 20;           // Tab-delimited text
  xlHTML = 44;                  // HTML format
  xlTemplate = 17;              // Excel template (.xlt)
  xlOpenXMLTemplate = 54;       // Excel 2007+ template (.xltx)
  xlAddIn = 18;                 // Excel add-in (.xla)
  xlOpenXMLAddIn = 55;          // Excel 2007+ add-in (.xlam)
  
  // Excel constants - Chart Types
  xlArea = 1;
  xlBar = 2;
  xlColumn = 3;
  xlLine = 4;
  xlPie = 5;
  xlRadar = -4151;
  xlXYScatter = -4169;
  xl3DArea = -4098;
  xl3DBar = -4099;
  xl3DColumn = -4100;
  xl3DLine = -4101;
  xl3DPie = -4102;
  xl3DSurface = -4103;
  
  // Excel constants - Sort Order
  xlAscending = 1;
  xlDescending = 2;
  
  // Excel constants - Format Condition Type
  xlCellValue = 1;
  xlExpression = 2;
  xlColorScale = 3;
  xlDataBar = 4;
  xlIconSet = 6;

  // Excel constants - Condition Operator
  xlBetween = 1;
  xlNotBetween = 2;
  xlEqual = 3;
  xlNotEqual = 4;
  xlGreater = 5;
  xlLess = 6;
  xlGreaterEqual = 7;
  xlLessEqual = 8;

implementation

{ Helper functions }

function ExcelChartTypeToOleEnum(ChartType: TDFExcelChartType): Integer;
begin
  case ChartType of
    ctArea: Result := xlArea;
    ctBar: Result := xlBar;
    ctColumn: Result := xlColumn;
    ctLine: Result := xlLine;
    ctPie: Result := xlPie;
    ctRadar: Result := xlRadar;
    ctXY: Result := xlXYScatter;
    ct3DArea: Result := xl3DArea;
    ct3DBar: Result := xl3DBar;
    ct3DColumn: Result := xl3DColumn;
    ct3DLine: Result := xl3DLine;
    ct3DPie: Result := xl3DPie;
    ct3DSurface: Result := xl3DSurface;
    else Result := xlColumn; // Default
  end;
end;

function OleEnumToExcelChartType(OleEnum: Integer): TDFExcelChartType;
begin
  case OleEnum of
    xlArea: Result := ctArea;
    xlBar: Result := ctBar;
    xlColumn: Result := ctColumn;
    xlLine: Result := ctLine;
    xlPie: Result := ctPie;
    xlRadar: Result := ctRadar;
    xlXYScatter: Result := ctXY;
    xl3DArea: Result := ct3DArea;
    xl3DBar: Result := ct3DBar;
    xl3DColumn: Result := ct3DColumn;
    xl3DLine: Result := ct3DLine;
    xl3DPie: Result := ct3DPie;
    xl3DSurface: Result := ct3DSurface;
    else Result := ctColumn; // Default
  end;
end;

function RGBToOleColor(Color: TColor): Integer;
begin
  Result := RGB(GetRValue(Color), GetGValue(Color), GetBValue(Color));
end;

function OleColorToRGB(OleColor: Integer): TColor;
begin
  Result := TColor(OleColor and $FFFFFF);
end;

{ TDFExcelFont implementation }

constructor TDFExcelFont.Create(Range: OleVariant);
begin
  inherited Create;
  FRange := Range;
end;

function TDFExcelFont.GetBold: Boolean;
begin
  Result := FRange.Font.Bold;
end;

procedure TDFExcelFont.SetBold(const Value: Boolean);
begin
  FRange.Font.Bold := Value;
end;

function TDFExcelFont.GetItalic: Boolean;
begin
  Result := FRange.Font.Italic;
end;

procedure TDFExcelFont.SetItalic(const Value: Boolean);
begin
  FRange.Font.Italic := Value;
end;

function TDFExcelFont.GetSize: Integer;
begin
  Result := FRange.Font.Size;
end;

procedure TDFExcelFont.SetSize(const Value: Integer);
begin
  FRange.Font.Size := Value;
end;

function TDFExcelFont.GetColor: TColor;
begin
  Result := OleColorToRGB(FRange.Font.Color);
end;

procedure TDFExcelFont.SetColor(const Value: TColor);
begin
  FRange.Font.Color := RGBToOleColor(Value);
end;

function TDFExcelFont.GetName: string;
begin
  Result := FRange.Font.Name;
end;

procedure TDFExcelFont.SetName(const Value: string);
begin
  FRange.Font.Name := Value;
end;

function TDFExcelFont.GetUnderline: Boolean;
begin
  Result := FRange.Font.Underline <> xlLineStyleNone;
end;

procedure TDFExcelFont.SetUnderline(const Value: Boolean);
begin
  if Value then
    FRange.Font.Underline := xlContinuous
  else
    FRange.Font.Underline := xlLineStyleNone;
end;

{ TDFExcelInterior implementation }

constructor TDFExcelInterior.Create(Range: OleVariant);
begin
  inherited Create;
  FRange := Range;
end;

function TDFExcelInterior.GetColor: TColor;
begin
  Result := OleColorToRGB(FRange.Interior.Color);
end;

procedure TDFExcelInterior.SetColor(const Value: TColor);
begin
  FRange.Interior.Color := RGBToOleColor(Value);
end;

function TDFExcelInterior.GetPattern: Integer;
begin
  Result := FRange.Interior.Pattern;
end;

procedure TDFExcelInterior.SetPattern(const Value: Integer);
begin
  FRange.Interior.Pattern := Value;
end;

{ TDFExcelBorders implementation }

constructor TDFExcelBorders.Create(Range: OleVariant);
begin
  inherited Create;
  FRange := Range;
end;

function TDFExcelBorders.GetLineStyle: Integer;
begin
  Result := FRange.Borders.LineStyle;
end;

procedure TDFExcelBorders.SetLineStyle(const Value: Integer);
begin
  FRange.Borders.LineStyle := Value;
end;

function TDFExcelBorders.GetWeight: Integer;
begin
  Result := FRange.Borders.Weight;
end;

procedure TDFExcelBorders.SetWeight(const Value: Integer);
begin
  FRange.Borders.Weight := Value;
end;

function TDFExcelBorders.GetColor: TColor;
begin
  Result := OleColorToRGB(FRange.Borders.Color);
end;

procedure TDFExcelBorders.SetColor(const Value: TColor);
begin
  FRange.Borders.Color := RGBToOleColor(Value);
end;

{ TDFExcelCell implementation }

constructor TDFExcelCell.Create(Cell: OleVariant);
begin
  inherited Create;
  FCell := Cell;
  FFont := TDFExcelFont.Create(Cell);
  FInterior := TDFExcelInterior.Create(Cell);
  FBorders := TDFExcelBorders.Create(Cell);
end;

destructor TDFExcelCell.Destroy;
begin
  FFont.Free;
  FInterior.Free;
  FBorders.Free;
  inherited;
end;

function TDFExcelCell.GetValue: Variant;
begin
  Result := FCell.Value;
end;

procedure TDFExcelCell.SetValue(const Value: Variant);
begin
  FCell.Value := Value;
end;

function TDFExcelCell.GetFormula: string;
begin
  Result := FCell.Formula;
end;

procedure TDFExcelCell.SetFormula(const Value: string);
begin
  FCell.Formula := Value;
end;

function TDFExcelCell.GetNumberFormat: string;
begin
  Result := FCell.NumberFormat;
end;

procedure TDFExcelCell.SetNumberFormat(const Value: string);
begin
  FCell.NumberFormat := Value;
end;

function TDFExcelCell.GetHorizontalAlignment: Integer;
begin
  Result := FCell.HorizontalAlignment;
end;

procedure TDFExcelCell.SetHorizontalAlignment(const Value: Integer);
begin
  FCell.HorizontalAlignment := Value;
end;

function TDFExcelCell.GetVerticalAlignment: Integer;
begin
  Result := FCell.VerticalAlignment;
end;

procedure TDFExcelCell.SetVerticalAlignment(const Value: Integer);
begin
  FCell.VerticalAlignment := Value;
end;

{ TDFExcelRange implementation }

constructor TDFExcelRange.Create(Range: OleVariant);
begin
  inherited Create;
  FRange := Range;
  FFont := TDFExcelFont.Create(Range);
  FInterior := TDFExcelInterior.Create(Range);
  FBorders := TDFExcelBorders.Create(Range);
end;

destructor TDFExcelRange.Destroy;
begin
  FFont.Free;
  FInterior.Free;
  FBorders.Free;
  inherited;
end;

function TDFExcelRange.GetValue: Variant;
begin
  Result := FRange.Value;
end;

procedure TDFExcelRange.SetValue(const Value: Variant);
begin
  FRange.Value := Value;
end;

function TDFExcelRange.GetFormula: string;
begin
  Result := FRange.Formula;
end;

procedure TDFExcelRange.SetFormula(const Value: string);
begin
  FRange.Formula := Value;
end;

function TDFExcelRange.GetHorizontalAlignment: Integer;
begin
  Result := FRange.HorizontalAlignment;
end;

procedure TDFExcelRange.SetHorizontalAlignment(const Value: Integer);
begin
  FRange.HorizontalAlignment := Value;
end;

function TDFExcelRange.GetVerticalAlignment: Integer;
begin
  Result := FRange.VerticalAlignment;
end;

procedure TDFExcelRange.SetVerticalAlignment(const Value: Integer);
begin
  FRange.VerticalAlignment := Value;
end;

procedure TDFExcelRange.Clear;
begin
  FRange.Clear;
end;

procedure TDFExcelRange.ClearContents;
begin
  FRange.ClearContents;
end;

procedure TDFExcelRange.ClearFormats;
begin
  FRange.ClearFormats;
end;

procedure TDFExcelRange.AutoFit;
begin
  FRange.Columns.AutoFit;
end;

procedure TDFExcelRange.Merge;
begin
  FRange.Merge;
end;

procedure TDFExcelRange.Unmerge;
begin
  FRange.UnMerge;
end;

procedure TDFExcelRange.Copy;
begin
  FRange.Copy;
end;

procedure TDFExcelRange.Cut;
begin
  FRange.Cut;
end;

procedure TDFExcelRange.Paste;
begin
  FRange.PasteSpecial;
end;

procedure TDFExcelRange.PasteSpecial(Format: Integer);
begin
  FRange.PasteSpecial(Format);
end;

procedure TDFExcelRange.Sort(Key1Range: string; Order1: Integer; Key2Range: string;
                            Order2: Integer; Header: Integer);
var
  Key1, Key2: OleVariant;
begin
  if Key1Range <> '' then
    Key1 := FRange.Worksheet.Range[Key1Range];
  
  if Key2Range <> '' then
    Key2 := FRange.Worksheet.Range[Key2Range]
  else
    Key2 := Unassigned;
  
  if Key2 = Unassigned then
    FRange.Sort(Key1, Order1, , , , , , Header)
  else
    FRange.Sort(Key1, Order1, Key2, , Order2, , , Header);
end;

function TDFExcelRange.CreateFormatCondition(ConditionType: Integer; Operator: Integer;
                                  Formula1: string; Formula2: string): OleVariant;
begin
  if Formula2 = '' then
    Result := FRange.FormatConditions.Add(ConditionType, Operator, Formula1)
  else
    Result := FRange.FormatConditions.Add(ConditionType, Operator, Formula1, Formula2);
end;

{ TDFExcelChart implementation }

constructor TDFExcelChart.Create(Chart: OleVariant);
begin
  inherited Create;
  FChart := Chart;
end;

function TDFExcelChart.GetChartType: TDFExcelChartType;
begin
  Result := OleEnumToExcelChartType(FChart.ChartType);
end;

procedure TDFExcelChart.SetChartType(const Value: TDFExcelChartType);
begin
  FChart.ChartType := ExcelChartTypeToOleEnum(Value);
end;

function TDFExcelChart.GetHasTitle: Boolean;
begin
  Result := FChart.HasTitle;
end;

procedure TDFExcelChart.SetHasTitle(const Value: Boolean);
begin
  FChart.HasTitle := Value;
end;

function TDFExcelChart.GetChartTitle: OleVariant;
begin
  Result := FChart.ChartTitle;
end;

function TDFExcelChart.GetHasLegend: Boolean;
begin
  Result := FChart.HasLegend;
end;

procedure TDFExcelChart.SetHasLegend(const Value: Boolean);
begin
  FChart.HasLegend := Value;
end;

procedure TDFExcelChart.Export(FileName: string; FilterName: string);
begin
  FChart.Export(FileName, FilterName);
end;

procedure TDFExcelChart.SetSourceData(Range: OleVariant; PlotBy: Integer);
begin
  FChart.SetSourceData(Range, PlotBy);
end;

{ TDFExcelWorksheet implementation }

constructor TDFExcelWorksheet.Create(Worksheet: OleVariant);
begin
  inherited Create;
  FWorksheet := Worksheet;
end;

destructor TDFExcelWorksheet.Destroy;
begin
  inherited;
end;

function TDFExcelWorksheet.GetName: string;
begin
  Result := FWorksheet.Name;
end;

procedure TDFExcelWorksheet.SetName(const Value: string);
begin
  FWorksheet.Name := Value;
end;

function TDFExcelWorksheet.GetCell(Row, Col: Integer): TDFExcelCell;
begin
  Result := TDFExcelCell.Create(FWorksheet.Cells[Row, Col]);
end;

function TDFExcelWorksheet.GetRange(const Address: string): TDFExcelRange;
begin
  Result := TDFExcelRange.Create(FWorksheet.Range[Address]);
end;

function TDFExcelWorksheet.GetUsedRange: TDFExcelRange;
begin
  Result := TDFExcelRange.Create(FWorksheet.UsedRange);
end;

function TDFExcelWorksheet.GetVisible: Boolean;
begin
  Result := FWorksheet.Visible;
end;

procedure TDFExcelWorksheet.SetVisible(const Value: Boolean);
begin
  FWorksheet.Visible := Value;
end;

function TDFExcelWorksheet.CopySheet(After: OleVariant): TDFExcelWorksheet;
begin
  FWorksheet.Copy(After);
  Result := TDFExcelWorksheet.Create(FWorksheet.Application.ActiveSheet);
end;

procedure TDFExcelWorksheet.Delete;
begin
  FWorksheet.Delete;
end;

procedure TDFExcelWorksheet.Activate;
begin
  FWorksheet.Activate;
end;

procedure TDFExcelWorksheet.InsertRow(RowIndex: Integer);
begin
  FWorksheet.Rows[RowIndex].Insert;
end;

procedure TDFExcelWorksheet.InsertColumn(ColIndex: Integer);
begin
  FWorksheet.Columns[ColIndex].Insert;
end;

procedure TDFExcelWorksheet.DeleteRow(RowIndex: Integer);
begin
  FWorksheet.Rows[RowIndex].Delete;
end;

procedure TDFExcelWorksheet.DeleteColumn(ColIndex: Integer);
begin
  FWorksheet.Columns[ColIndex].Delete;
end;

function TDFExcelWorksheet.CreateChart(ChartType: TDFExcelChartType; DataRange: string): TDFExcelChart;
var
  ChartObj: OleVariant;
  Range: OleVariant;
begin
  Range := FWorksheet.Range[DataRange];
  ChartObj := FWorksheet.ChartObjects.Add(100, 100, 400, 300);
  ChartObj.Chart.SetSourceData(Range);
  ChartObj.Chart.ChartType := ExcelChartTypeToOleEnum(ChartType);
  Result := TDFExcelChart.Create(ChartObj.Chart);
end;

procedure TDFExcelWorksheet.PrintOut(From, To_, Copies: Integer);
begin
  if (From = 0) and (To_ = 0) then
    FWorksheet.PrintOut(, , Copies)
  else
    FWorksheet.PrintOut(From, To_, Copies);
end;

procedure TDFExcelWorksheet.PrintPreview;
begin
  FWorksheet.PrintPreview;
end;

{ TDFExcelWorkbook implementation }

constructor TDFExcelWorkbook.Create(Workbook: OleVariant);
begin
  inherited Create;
  FWorkbook := Workbook;
  FSheets := TList.Create;
end;

destructor TDFExcelWorkbook.Destroy;
var
  i: Integer;
begin
  for i := 0 to FSheets.Count - 1 do
    TDFExcelWorksheet(FSheets[i]).Free;
  FSheets.Free;
  inherited;
end;

function TDFExcelWorkbook.GetWorksheet(Index: Integer): TDFExcelWorksheet;
var
  Sheet: TDFExcelWorksheet;
begin
  if (Index < 1) or (Index > FWorkbook.Worksheets.Count) then
    raise EDFExcelException.CreateFmt('Worksheet index out of bounds: %d', [Index]);

  // Check if we already have this worksheet in our list
  if (Index <= FSheets.Count) and (FSheets[Index - 1] <> nil) then
    Result := TDFExcelWorksheet(FSheets[Index - 1])
  else
  begin
    // Create new worksheet object
    Sheet := TDFExcelWorksheet.Create(FWorkbook.Worksheets[Index]);
    
    // Ensure we have enough slots in the list
    while FSheets.Count < Index do
      FSheets.Add(nil);
      
    // Store the worksheet in our list
    FSheets[Index - 1] := Sheet;
    Result := Sheet;
  end;
end;

function TDFExcelWorkbook.GetWorksheetCount: Integer;
begin
  Result := FWorkbook.Worksheets.Count;
end;

function TDFExcelWorkbook.GetFileName: string;
begin
  Result := FWorkbook.FullName;
end;

function TDFExcelWorkbook.AddWorksheet(const Name: string): TDFExcelWorksheet;
var
  Sheet: OleVariant;
begin
  if Name = '' then
    Sheet := FWorkbook.Worksheets.Add
  else
  begin
    Sheet := FWorkbook.Worksheets.Add;
    Sheet.Name := Name;
  end;
  
  Result := TDFExcelWorksheet.Create(Sheet);
  FSheets.Add(Result);
end;

function TDFExcelWorkbook.WorksheetByName(const Name: string): TDFExcelWorksheet;
var
  i: Integer;
  Sheet: OleVariant;
begin
  Result := nil;
  
  // First check our cache
  for i := 0 to FSheets.Count - 1 do
  begin
    if (FSheets[i] <> nil) and 
       (TDFExcelWorksheet(FSheets[i]).Name = Name) then
    begin
      Result := TDFExcelWorksheet(FSheets[i]);
      Exit;
    end;
  end;
  
  // Not found in cache, try to find by name
  try
    Sheet := FWorkbook.Worksheets[Name];
    Result := TDFExcelWorksheet.Create(Sheet);
    FSheets.Add(Result);
  except
    on E: Exception do
      raise EDFExcelException.CreateFmt('Worksheet "%s" not found', [Name]);
  end;
end;

procedure TDFExcelWorkbook.Save;
begin
  FWorkbook.Save;
end;

procedure TDFExcelWorkbook.SaveAs(const FileName: string; FileFormat: Integer);
begin
  FWorkbook.SaveAs(FileName, FileFormat);
end;

procedure TDFExcelWorkbook.Close(SaveChanges: Boolean);
begin
  if SaveChanges then
    FWorkbook.Close(True)
  else
    FWorkbook.Close(False);
end;

{ TDFExcelApplication implementation }

constructor TDFExcelApplication.Create(Visible: Boolean);
begin
  inherited Create;
  FWorkbooks := TList.Create;
  
  try
    // Create Excel application
    FExcel := CreateOleObject('Excel.Application');
    FExcel.Visible := Visible;
    
    // Disable alerts
    FExcel.DisplayAlerts := False;
  except
    on E: Exception do
      raise EDFExcelException.Create('Failed to create Excel application: ' + E.Message);
  end;
end;

destructor TDFExcelApplication.Destroy;
var
  i: Integer;
begin
  // Free all workbooks
  for i := 0 to FWorkbooks.Count - 1 do
    TDFExcelWorkbook(FWorkbooks[i]).Free;
  FWorkbooks.Free;
  
  // Quit Excel if it's still running
  if not VarIsEmpty(FExcel) then
  begin
    try
      FExcel.Quit;
    except
      // Ignore errors during cleanup
    end;
  end;
  
  inherited;
end;

function TDFExcelApplication.GetVisible: Boolean;
begin
  Result := FExcel.Visible;
end;

procedure TDFExcelApplication.SetVisible(const Value: Boolean);
begin
  FExcel.Visible := Value;
end;

function TDFExcelApplication.GetActiveWorkbook: TDFExcelWorkbook;
var
  i: Integer;
  ActiveName: string;
begin
  Result := nil;
  
  if VarIsEmpty(FExcel.ActiveWorkbook) then
    Exit;
    
  ActiveName := FExcel.ActiveWorkbook.FullName;
  
  // Find the workbook in our list
  for i := 0 to FWorkbooks.Count - 1 do
  begin
    if TDFExcelWorkbook(FWorkbooks[i]).FileName = ActiveName then
    begin
      Result := TDFExcelWorkbook(FWorkbooks[i]);
      Exit;
    end;
  end;
  
  // If not found, create a new wrapper for it
  Result := TDFExcelWorkbook.Create(FExcel.ActiveWorkbook);
  FWorkbooks.Add(Result);
end;

function TDFExcelApplication.GetWorkbook(Index: Integer): TDFExcelWorkbook;
var
  Workbook: TDFExcelWorkbook;
begin
  if (Index < 1) or (Index > FExcel.Workbooks.Count) then
    raise EDFExcelException.CreateFmt('Workbook index out of bounds: %d', [Index]);

  // Check if we already have this workbook in our list
  if (Index <= FWorkbooks.Count) and (FWorkbooks[Index - 1] <> nil) then
    Result := TDFExcelWorkbook(FWorkbooks[Index - 1])
  else
  begin
    // Create new workbook object
    Workbook := TDFExcelWorkbook.Create(FExcel.Workbooks[Index]);
    
    // Ensure we have enough slots in the list
    while FWorkbooks.Count < Index do
      FWorkbooks.Add(nil);
      
    // Store the workbook in our list
    FWorkbooks[Index - 1] := Workbook;
    Result := Workbook;
  end;
end;

function TDFExcelApplication.GetWorkbookCount: Integer;
begin
  Result := FExcel.Workbooks.Count;
end;

function TDFExcelApplication.CreateWorkbook: TDFExcelWorkbook;
var
  Workbook: TDFExcelWorkbook;
begin
  Workbook := TDFExcelWorkbook.Create(FExcel.Workbooks.Add);
  FWorkbooks.Add(Workbook);
  Result := Workbook;
end;

function TDFExcelApplication.OpenWorkbook(const FileName: string): TDFExcelWorkbook;
var
  Workbook: TDFExcelWorkbook;
begin
  try
    Workbook := TDFExcelWorkbook.Create(FExcel.Workbooks.Open(FileName));
    FWorkbooks.Add(Workbook);
    Result := Workbook;
  except
    on E: Exception do
      raise EDFExcelException.CreateFmt('Failed to open workbook "%s": %s', [FileName, E.Message]);
  end;
end;

procedure TDFExcelApplication.Quit;
begin
  if not VarIsEmpty(FExcel) then
  begin
    try
      FExcel.Quit;
      FExcel := Unassigned;
    except
      on E: Exception do
        raise EDFExcelException.Create('Failed to quit Excel: ' + E.Message);
    end;
  end;
end;

procedure TDFExcelApplication.DisplayAlerts(Value: Boolean);
begin
  FExcel.DisplayAlerts := Value;
end;

procedure TDFExcelApplication.ScreenUpdating(Value: Boolean);
begin
  FExcel.ScreenUpdating := Value;
end;

procedure TDFExcelApplication.Calculate;
begin
  FExcel.Calculate;
end;

function TDFExcelApplication.Evaluate(const Name: string): Variant;
begin
  Result := FExcel.Evaluate(Name);
end;

end.