# Open XML SDK Snippets and Item Templates

A collection of snippets for creation and manipulation of Open XML documents, spreadsheets and presentations.

- [Usage](#usage)
- [Snippets](#snippets)
    - [Excel](#excel)
        - [Excel_CreateSpreadsheet](#excel_createspreadsheet)
        - [Excel_GetAllSheets](#excel_getallsheets)
        - [Excel_GetSheet](#excel_getsheet)
        - [Excel_InsertTextIntoCell](#excel_inserttextintocell)
    - [PowerPoint](#powerpoint)
        - [PowerPoint_CreatePresentation](#powerpoint_createpresentation)
    - [Word](#word)
        - [Word_CreateDocument](#word_createdocument)
- [Item Templates](#item-templates)
    - [Open XML SDK Utils](#open-xml-sdk-utils)
    - [Open XML SDK PowerPoint Utils](#open-xml-sdk-powerpoint-utils)

---

## Usage

This extension consists of 2 parts: snippets that can be inserted directly in code and item templates with utility methods to facilitate working on Office Open XML files with the Open XML SDK.

The item templates are added by right-clicking the project in Solution Explorer | Add | New Item... | Open XML SDK \<PowerPoint | Excel\> Utils. The item templates contain code that is too large to fit in a snippet, but could be useful to developers. Some of the snippets use the item templates, but the utility methods can be used directly without the snippets.

Snippets are added to IntelliSense, to use them begin typing the name of the snippet you want to insert and you will see a list of available snippets. Some of the snippets rely on your project containing an instance of the utility methods. If a snippet requires utility methods, it is noted in comments in the snippet.

*The snippets and the item templates both require the [Open XML SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml/) be installed from nuget.*

## Snippets
### Excel

#### Excel_CreateSpreadsheet

```csharp
// Create a workbook document by supplying the file path.
// By default, AutoSave = true, Editable = true, and Type = xlsx.
SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(${1}, SpreadsheetDocumentType.Workbook);

// Add a WorkbookPart to the document.
WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
workbookpart.Workbook = new Workbook();

// Add a WorksheetPart to the WorkbookPart.
WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
worksheetPart.Worksheet = new Worksheet(new SheetData());

// Add Sheets to the workbook.
Sheets? sheets = spreadsheetDocument?.WorkbookPart?.Workbook.AppendChild<Sheets>(new Sheets());

// Append a new worksheet and associate it with the workbook.
Sheet? sheet = new Sheet() { Id = spreadsheetDocument?.WorkbookPart?.GetIdOfPart(worksheetPart), SheetId = 1, Name = ${2} };

if (sheets != null && sheet != null)
{
    sheets.Append(sheet);
}
```

#### Excel_GetAllSheets

```csharp
Sheets? theSheets = null;

using (SpreadsheetDocument? document =
    SpreadsheetDocument.Open(${1}, false))
{
    WorkbookPart? wbPart = document.WorkbookPart;
    theSheets = wbPart?.Workbook.Sheets;
}
```

#### Excel_GetSheet

```csharp
Sheet? sheet = null;

using (SpreadsheetDocument? document = SpreadsheetDocument.Open($FilePath$, true))
{
    IEnumerable<Sheet>? sheets = document?.WorkbookPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == $SheetName$);

    if (sheets != null && sheets.Count() > 0)
    {
        sheet = sheets.FirstOrDefault();
    }
}
```

#### Excel_InsertTextIntoCell

*REQUIRED: Add Open XML SDK Excel Utils item. Right-click the project in Solution Explorer | Add | New Item... | Open XML SDK Excel Utils*

```csharp
// REQUIRED: Add Open XML SDK Excel Utils item. Right-click the project in Solution Explorer | Add | New Item... | Open XML SDK Excel Utils
SpreadsheetDocument spreadsheetDocument = ExcelUtils.InsertText($SpreadSheetDocument$, $Content$, $Row$, $Column$);
// TODO: manipulate the spreadsheet
spreadsheetDocument.Save();
spreadsheetDocument.Close();
```

### PowerPoint

#### PowerPoint_CreatePresentation

*REQUIRED: Add Open XML SDK PowerPoint Utils item. Right-click the project in Solution Explorer | Add | New Item... | Open XML SDK PowerPoint Utils*
```csharp
// REQUIRED: Add Open XML SDK PowerPoint Utils item. Right-click the project in Solution Explorer | Add | New Item... | Open XML SDK PowerPoint Utils
PresentationDocument presentationDocument = PowerPointUtils.CreatePresentation(${1});
// TODO: manipulate the presentation
presentationDocument.Save();
presentationDocument.Close();
```

### Word

#### Word_CreateDocument

```csharp
// Create a document by supplying the filepath.
using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(${1}, WordprocessingDocumentType.Document))
{
	// Add a main document part.
	MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

	// Create the document structure and add some text.
	mainPart.Document = new Document();
	Body body = mainPart.Document.AppendChild(new Body());
	Paragraph para = body.AppendChild(new Paragraph());
	Run run = para.AppendChild(new Run());
	run.AppendChild(new Text(${2}));
}
```

## Item Templates

### Open XML SDK Utils

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlSdkUtils
{
    public class ExcelUtils
    {
        // Given a document name and text,
        // inserts a new worksheet and writes the text into a cell of the new worksheet.

        public static SpreadsheetDocument InsertText(SpreadsheetDocument spreadSheet, string text, string columnName, uint rowIndex, WorksheetPart? wsp)
        {
            if (string.IsNullOrEmpty(text))
            {
                throw new ArgumentNullException(nameof(text));
            }
            if (string.IsNullOrEmpty(columnName))
            {
                throw new ArgumentNullException(nameof(columnName));
            }
            if (string.IsNullOrEmpty(rowIndex.ToString()))
            {
                throw new ArgumentNullException(nameof(rowIndex));
            }

            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (spreadSheet != null && spreadSheet.WorkbookPart != null)
            {
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Any())
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
            }
            else
            {
                throw new ArgumentNullException(nameof(spreadSheet));
            }

            // Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem(text, shareStringPart);

            WorksheetPart worksheetPart;
            // Insert a new worksheet if wsp is null.
            if (wsp == null)
            {
                worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);
            }
            else
            {
                worksheetPart = wsp;
            }

            // Insert cell into the worksheet.
            Cell cell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart);

            // Set the value of the cell.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            worksheetPart.Worksheet.Save();

            if (spreadSheet != null)
            {
                return spreadSheet;
            }
            else
            {
                throw new ArgumentNullException(nameof(spreadSheet));
            }
        }

        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        public static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        // Given a WorkbookPart, inserts a new worksheet.
        public static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new WorksheetPart to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            if (workbookPart != null && workbookPart.Workbook != null)
            {
                Sheets? sheets = workbookPart.Workbook.GetFirstChild<Sheets>();

                if (sheets == null)
                {
                    sheets = new Sheets();
                    workbookPart.Workbook.AddChild(sheets);
                }
                string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

                // Get a unique ID for the new sheet.
                uint sheetId = 1;

                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    IEnumerable<Sheet> sheetElements = sheets.Elements<Sheet>();

                    foreach (Sheet el in sheetElements)
                    {
                        if (el != null && !string.IsNullOrEmpty(el.SheetId))
                        {
                            sheetId += el.SheetId.Value;
                        }
                    }
                }

                string sheetName = string.Concat("Sheet", sheetId);

                // Append the new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);
                workbookPart.Workbook.Save();

                return newWorksheetPart;
            }
            else
            {
                throw new ArgumentNullException(nameof(workbookPart));
            }
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet.
        // If the cell already exists, returns it.
        public static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            if (worksheetPart != null)
            {
                Worksheet worksheet = worksheetPart.Worksheet ?? new Worksheet();

                SheetData? sheetData = worksheet.GetFirstChild<SheetData>();

                if (sheetData == null)
                {
                    sheetData = new SheetData();
                    worksheet.AddChild(sheetData);
                }

                string cellReference = columnName + rowIndex;

                // If the worksheet does not contain a row with the specified row index, insert one.
                Row row;
                if (sheetData.Elements<Row>().Where(r => !string.IsNullOrEmpty(r.RowIndex) && r.RowIndex == rowIndex).Count() != 0)
                {
                    row = sheetData.Elements<Row>().Where(r => !string.IsNullOrEmpty(r.RowIndex) && r.RowIndex == rowIndex).First();
                }
                else
                {
                    row = new Row() { RowIndex = rowIndex };
                    sheetData.Append(row);
                }

                // If a cell with the specified column name doesn’t exist, insert one.
                if (row.Elements<Cell>().Where(c => !string.IsNullOrEmpty(c.CellReference) && c.CellReference.Value == columnName + rowIndex).Count() > 0)
                {
                    return row.Elements<Cell>().Where(c => !string.IsNullOrEmpty(c.CellReference) && c.CellReference.Value == cellReference).First();
                }
                else
                {
                    // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                    Cell? refCell = null;

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        if (!string.IsNullOrEmpty(cell.CellReference) && string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }

                    if (refCell != null)
                    {
                        Cell newCell = new Cell() { CellReference = cellReference };
                        row.InsertBefore(newCell, refCell);

                        worksheet.Save();
                        return newCell;
                    }
                    else
                    {
                        throw new InvalidDataException();
                    }
                }
            }
            else
            {
                throw new ArgumentNullException(nameof(worksheetPart));
            }
        }

        public static Sheet? GetSheet(SpreadsheetDocument spreadsheetDocument, string sheetName)
        {
            IEnumerable<Sheet>? sheets = spreadsheetDocument.WorkbookPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);

            if (sheets != null && sheets.Count() > 0)
            {
                return sheets.FirstOrDefault();
            }

            return null;
        }
    }
}
```

### Open XML SDK PowerPoint Utils

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXmlSdkUtils
{
    public class PowerPointUtils
    {
        public static PresentationDocument CreatePresentation(string filepath)
        {
            // Create a presentation at a specified file path. The presentation document type is pptx by default.
            PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            _createPresentationParts(presentationPart);

            return presentationDoc;
        }

        private static void _createPresentationParts(PresentationPart presentationPart)
        {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
            SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

            SlidePart slidePart1;
            SlideLayoutPart slideLayoutPart1;
            SlideMasterPart slideMasterPart1;
            ThemePart themePart1;


            slidePart1 = _createSlidePart(presentationPart);
            slideLayoutPart1 = _createSlideLayoutPart(slidePart1);
            slideMasterPart1 = _createSlideMasterPart(slideLayoutPart1);
            themePart1 = _createTheme(slideMasterPart1);

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
            presentationPart.AddPart(slideMasterPart1, "rId1");
            presentationPart.AddPart(themePart1, "rId5");
        }

        private static SlidePart _createSlidePart(PresentationPart presentationPart)
        {
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
            slidePart1.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup()),
                            new P.Shape(
                                new P.NonVisualShapeProperties(
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                new P.ShapeProperties(),
                                new P.TextBody(
                                    new BodyProperties(),
                                    new ListStyle(),
                                    new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
                    new ColorMapOverride(new MasterColorMapping()));
            return slidePart1;
        }

        private static SlideLayoutPart _createSlideLayoutPart(SlidePart slidePart1)
        {
            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            SlideLayout slideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(new EndParagraphRunProperties()))))),
            new ColorMapOverride(new MasterColorMapping()));
            slideLayoutPart1.SlideLayout = slideLayout;
            return slideLayoutPart1;
        }

        private static SlideMasterPart _createSlideMasterPart(SlideLayoutPart slideLayoutPart1)
        {
            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
            SlideMaster slideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph())))),
            new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
            new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
            new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
            slideMasterPart1.SlideMaster = slideMaster;

            return slideMasterPart1;
        }

        private static ThemePart _createTheme(SlideMasterPart slideMasterPart1)
        {
            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
            D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

            D.ThemeElements themeElements1 = new D.ThemeElements(
            new D.ColorScheme(
              new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
              new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
              new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
              new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
              new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
              new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
              new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
              new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
              new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
              new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
              new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
              new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" }))
            { Name = "Office" },
              new D.FontScheme(
              new D.MajorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }),
              new D.MinorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }))
              { Name = "Office" },
              new D.FormatScheme(
              new D.FillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                  new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                 new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 35000 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                 new D.SaturationModulation() { Val = 350000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 100000 }
                ),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.NoFill(),
              new D.PatternFill(),
              new D.GroupFill()),
              new D.LineStyleList(
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              }),
              new D.EffectStyleList(
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
              new D.BackgroundFillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true })))
              { Name = "Office" });

            theme1.Append(themeElements1);
            theme1.Append(new D.ObjectDefaults());
            theme1.Append(new D.ExtraColorSchemeList());

            themePart1.Theme = theme1;
            return themePart1;

        }
    }
}
```

## Release Notes

Please see the [Changelog](CHANGELOG.md)

---
<p align="right">Logo created with <a href="https://logomakr.com/">LogoMakr.com</a></p>
