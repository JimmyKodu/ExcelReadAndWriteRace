using System.Data;
using System.Diagnostics;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using MiniExcelLibs;
using OfficeOpenXml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

Console.OutputEncoding = Encoding.UTF8;
Console.WriteLine("=".PadRight(80, '='));
Console.WriteLine("Excel Libraries Performance Comparison");
Console.WriteLine("=".PadRight(80, '='));
Console.WriteLine();

var benchmarks = new ExcelBenchmark();
await benchmarks.RunAllBenchmarksAsync();

public class ExcelBenchmark
{
    private const int RowCount = 10000;
    private const int ColumnCount = 10;
    private readonly string testDataPath = Path.Combine(Path.GetTempPath(), "ExcelBenchmark");
    
    public async Task RunAllBenchmarksAsync()
    {
        // Ensure directory exists and is clean
        if (Directory.Exists(testDataPath))
        {
            Directory.Delete(testDataPath, true);
        }
        Directory.CreateDirectory(testDataPath);
        
        Console.WriteLine($"Test Configuration:");
        Console.WriteLine($"  Rows: {RowCount:N0}");
        Console.WriteLine($"  Columns: {ColumnCount}");
        Console.WriteLine($"  Test Directory: {testDataPath}");
        Console.WriteLine();
        
        // Write Tests - XLSX
        Console.WriteLine("=" + "=".PadRight(78, '='));
        Console.WriteLine("WRITE PERFORMANCE TESTS (XLSX Format)");
        Console.WriteLine("=" + "=".PadRight(78, '='));
        Console.WriteLine();
        
        var miniExcelWriteTime = await BenchmarkWriteMiniExcelAsync();
        var excelDataReaderWriteTime = "N/A (Read-only)";
        var openXmlWriteTime = await BenchmarkWriteOpenXmlAsync();
        var closedXmlWriteTime = await BenchmarkWriteClosedXMLAsync();
        var epplusWriteTime = await BenchmarkWriteEPPlusAsync();
        
        Console.WriteLine();
        
        // Read Tests - XLSX
        Console.WriteLine("=" + "=".PadRight(78, '='));
        Console.WriteLine("READ PERFORMANCE TESTS (XLSX Format)");
        Console.WriteLine("=" + "=".PadRight(78, '='));
        Console.WriteLine();
        
        var miniExcelReadTime = await BenchmarkReadMiniExcelAsync();
        var excelDataReaderReadTime = await BenchmarkReadExcelDataReaderAsync();
        var openXmlReadTime = await BenchmarkReadOpenXmlAsync();
        var closedXmlReadTime = await BenchmarkReadClosedXMLAsync();
        var epplusReadTime = await BenchmarkReadEPPlusAsync();
        
        Console.WriteLine();
        
        // XLS Format Compatibility Tests
        Console.WriteLine("=" + "=".PadRight(78, '='));
        Console.WriteLine("XLS FORMAT (Excel 97-2003) COMPATIBILITY TESTS");
        Console.WriteLine("=" + "=".PadRight(78, '='));
        Console.WriteLine();
        
        var npoiXlsWriteTime = await BenchmarkWriteNpoiXlsAsync();
        var excelDataReaderXlsReadTime = await BenchmarkReadExcelDataReaderXlsAsync();
        var miniExcelXlsReadTime = await BenchmarkReadMiniExcelXlsAsync();
        
        Console.WriteLine();
        Console.WriteLine("Note: Only ExcelDataReader and NPOI support the legacy XLS format.");
        Console.WriteLine("      DocumentFormat.OpenXml, ClosedXML, EPPlus, and MiniExcel only support XLSX.");
        Console.WriteLine();
        
        // Summary
        Console.WriteLine("=" + "=".PadRight(78, '='));
        Console.WriteLine("SUMMARY - XLSX Format");
        Console.WriteLine("=" + "=".PadRight(78, '='));
        Console.WriteLine();
        
        Console.WriteLine("┌─────────────────────────────────────┬─────────────────┬─────────────────┐");
        Console.WriteLine("│ Library                             │ Write Time      │ Read Time       │");
        Console.WriteLine("├─────────────────────────────────────┼─────────────────┼─────────────────┤");
        Console.WriteLine($"│ MiniExcel 1.41.4                    │ {miniExcelWriteTime,-15} │ {miniExcelReadTime,-15} │");
        Console.WriteLine($"│ ExcelDataReader 3.8.0               │ {excelDataReaderWriteTime,-15} │ {excelDataReaderReadTime,-15} │");
        Console.WriteLine($"│ DocumentFormat.OpenXml 3.3.0        │ {openXmlWriteTime,-15} │ {openXmlReadTime,-15} │");
        Console.WriteLine($"│ ClosedXML 0.105.0                   │ {closedXmlWriteTime,-15} │ {closedXmlReadTime,-15} │");
        Console.WriteLine($"│ EPPlus 4.5.3.3                      │ {epplusWriteTime,-15} │ {epplusReadTime,-15} │");
        Console.WriteLine("└─────────────────────────────────────┴─────────────────┴─────────────────┘");
        Console.WriteLine();
        
        Console.WriteLine("=" + "=".PadRight(78, '='));
        Console.WriteLine("SUMMARY - XLS Format (Excel 97-2003)");
        Console.WriteLine("=" + "=".PadRight(78, '='));
        Console.WriteLine();
        
        Console.WriteLine("┌─────────────────────────────────────┬─────────────────┬─────────────────┐");
        Console.WriteLine("│ Library                             │ Write Time      │ Read Time       │");
        Console.WriteLine("├─────────────────────────────────────┼─────────────────┼─────────────────┤");
        Console.WriteLine($"│ MiniExcel 1.41.4                    │ {"N/A",-15} │ {miniExcelXlsReadTime,-15} │");
        Console.WriteLine($"│ ExcelDataReader 3.8.0               │ {excelDataReaderWriteTime,-15} │ {excelDataReaderXlsReadTime,-15} │");
        Console.WriteLine($"│ DocumentFormat.OpenXml 3.3.0        │ {"N/A",-15} │ {"N/A",-15} │");
        Console.WriteLine($"│ ClosedXML 0.105.0                   │ {"N/A",-15} │ {"N/A",-15} │");
        Console.WriteLine($"│ EPPlus 4.5.3.3                      │ {"N/A",-15} │ {"N/A",-15} │");
        Console.WriteLine($"│ NPOI 2.7.3                          │ {npoiXlsWriteTime,-15} │ {"N/A",-15} │");
        Console.WriteLine("└─────────────────────────────────────┴─────────────────┴─────────────────┘");
        Console.WriteLine();
    }
    
    private List<Dictionary<string, object>> GenerateTestData()
    {
        var data = new List<Dictionary<string, object>>();
        var baseDate = new DateTime(2024, 1, 1);
        
        // Data rows
        for (int i = 1; i <= RowCount; i++)
        {
            var row = new Dictionary<string, object>
            {
                { "Column1", i },
                { "Column2", $"Name{i}" },
                { "Column3", i * 100.50 },
                { "Column4", baseDate.AddDays(i) },
                { "Column5", i % 2 == 0 },
                { "Column6", $"Email{i}@example.com" },
                { "Column7", i * 1.5 },
                { "Column8", $"Address {i}" },
                { "Column9", i % 100 },
                { "Column10", $"Notes for row {i}" }
            };
            data.Add(row);
        }
        
        return data;
    }
    
    private async Task<string> BenchmarkWriteMiniExcelAsync()
    {
        var filePath = Path.Combine(testDataPath, "MiniExcel_Write.xlsx");
        var data = GenerateTestData();
        
        var sw = Stopwatch.StartNew();
        await MiniExcel.SaveAsAsync(filePath, data, overwriteFile: true);
        sw.Stop();
        
        var fileSize = new FileInfo(filePath).Length;
        Console.WriteLine($"✓ MiniExcel Write:          {sw.ElapsedMilliseconds,6} ms (File size: {fileSize / 1024:N0} KB)");
        return $"{sw.ElapsedMilliseconds} ms";
    }
    
    private Task<string> BenchmarkWriteOpenXmlAsync()
    {
        var filePath = Path.Combine(testDataPath, "OpenXml_Write.xlsx");
        var data = GenerateTestData();
        
        var sw = Stopwatch.StartNew();
        
        using (var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            };
            sheets.Append(sheet);
            
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            
            // Write header
            var headerRow = new Row { RowIndex = 1 };
            var columnNames = data.First().Keys.ToList();
            for (int colIndex = 0; colIndex < columnNames.Count; colIndex++)
            {
                var cell = new Cell
                {
                    CellReference = GetColumnName(colIndex + 1) + "1",
                    DataType = CellValues.String,
                    CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(columnNames[colIndex])
                };
                headerRow.Append(cell);
            }
            sheetData!.Append(headerRow);
            
            // Write data
            uint rowIndex = 2;
            foreach (var dataRow in data)
            {
                var row = new Row { RowIndex = rowIndex++ };
                
                int colIndex = 0;
                foreach (var kvp in dataRow)
                {
                    colIndex++;
                    var cell = new Cell
                    {
                        CellReference = GetColumnName(colIndex) + (rowIndex - 1),
                        DataType = CellValues.String,
                        CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(kvp.Value?.ToString() ?? "")
                    };
                    row.Append(cell);
                }
                
                sheetData!.Append(row);
            }
        }
        
        sw.Stop();
        
        var fileSize = new FileInfo(filePath).Length;
        Console.WriteLine($"✓ OpenXml Write:            {sw.ElapsedMilliseconds,6} ms (File size: {fileSize / 1024:N0} KB)");
        return Task.FromResult($"{sw.ElapsedMilliseconds} ms");
    }
    
    private Task<string> BenchmarkWriteClosedXMLAsync()
    {
        var filePath = Path.Combine(testDataPath, "ClosedXML_Write.xlsx");
        var data = GenerateTestData();
        
        var sw = Stopwatch.StartNew();
        
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Sheet1");
            
            // Write header
            var columnNames = data.First().Keys.ToList();
            for (int colIndex = 0; colIndex < columnNames.Count; colIndex++)
            {
                worksheet.Cell(1, colIndex + 1).Value = columnNames[colIndex];
            }
            
            // Write data
            int rowIndex = 2;
            foreach (var dataRow in data)
            {
                int colIndex = 1;
                foreach (var kvp in dataRow)
                {
                    worksheet.Cell(rowIndex, colIndex++).Value = kvp.Value?.ToString() ?? "";
                }
                rowIndex++;
            }
            
            workbook.SaveAs(filePath);
        }
        
        sw.Stop();
        
        var fileSize = new FileInfo(filePath).Length;
        Console.WriteLine($"✓ ClosedXML Write:          {sw.ElapsedMilliseconds,6} ms (File size: {fileSize / 1024:N0} KB)");
        return Task.FromResult($"{sw.ElapsedMilliseconds} ms");
    }
    
    private Task<string> BenchmarkWriteEPPlusAsync()
    {
        var filePath = Path.Combine(testDataPath, "EPPlus_Write.xlsx");
        var data = GenerateTestData();
        
        var sw = Stopwatch.StartNew();
        
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");
            
            // Write header
            var columnNames = data.First().Keys.ToList();
            for (int colIndex = 0; colIndex < columnNames.Count; colIndex++)
            {
                worksheet.Cells[1, colIndex + 1].Value = columnNames[colIndex];
            }
            
            // Write data
            int rowIndex = 2;
            foreach (var dataRow in data)
            {
                int colIndex = 1;
                foreach (var kvp in dataRow)
                {
                    worksheet.Cells[rowIndex, colIndex++].Value = kvp.Value?.ToString() ?? "";
                }
                rowIndex++;
            }
            
            package.SaveAs(new FileInfo(filePath));
        }
        
        sw.Stop();
        
        var fileSize = new FileInfo(filePath).Length;
        Console.WriteLine($"✓ EPPlus Write:             {sw.ElapsedMilliseconds,6} ms (File size: {fileSize / 1024:N0} KB)");
        return Task.FromResult($"{sw.ElapsedMilliseconds} ms");
    }
    
    private async Task<string> BenchmarkReadMiniExcelAsync()
    {
        var filePath = Path.Combine(testDataPath, "MiniExcel_Write.xlsx");
        
        var sw = Stopwatch.StartNew();
        var rows = await MiniExcel.QueryAsync(filePath);
        var count = rows.Count();
        sw.Stop();
        
        Console.WriteLine($"✓ MiniExcel Read:           {sw.ElapsedMilliseconds,6} ms (Rows read: {count:N0})");
        return $"{sw.ElapsedMilliseconds} ms";
    }
    
    private Task<string> BenchmarkReadExcelDataReaderAsync()
    {
        var filePath = Path.Combine(testDataPath, "MiniExcel_Write.xlsx");
        
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        
        var sw = Stopwatch.StartNew();
        
        int count = 0;
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet();
                if (result.Tables.Count > 0)
                {
                    count = result.Tables[0].Rows.Count;
                }
            }
        }
        
        sw.Stop();
        
        Console.WriteLine($"✓ ExcelDataReader Read:     {sw.ElapsedMilliseconds,6} ms (Rows read: {count:N0})");
        return Task.FromResult($"{sw.ElapsedMilliseconds} ms");
    }
    
    private Task<string> BenchmarkReadOpenXmlAsync()
    {
        var filePath = Path.Combine(testDataPath, "OpenXml_Write.xlsx");
        
        var sw = Stopwatch.StartNew();
        
        int count = 0;
        using (var document = SpreadsheetDocument.Open(filePath, false))
        {
            var workbookPart = document.WorkbookPart;
            var worksheetPart = workbookPart!.WorksheetParts.First();
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            count = sheetData.Elements<Row>().Count();
        }
        
        sw.Stop();
        
        Console.WriteLine($"✓ OpenXml Read:             {sw.ElapsedMilliseconds,6} ms (Rows read: {count:N0})");
        return Task.FromResult($"{sw.ElapsedMilliseconds} ms");
    }
    
    private Task<string> BenchmarkReadClosedXMLAsync()
    {
        var filePath = Path.Combine(testDataPath, "ClosedXML_Write.xlsx");
        
        var sw = Stopwatch.StartNew();
        
        int count = 0;
        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet(1);
            count = worksheet.RowsUsed().Count();
        }
        
        sw.Stop();
        
        Console.WriteLine($"✓ ClosedXML Read:           {sw.ElapsedMilliseconds,6} ms (Rows read: {count:N0})");
        return Task.FromResult($"{sw.ElapsedMilliseconds} ms");
    }
    
    private Task<string> BenchmarkReadEPPlusAsync()
    {
        var filePath = Path.Combine(testDataPath, "EPPlus_Write.xlsx");
        
        var sw = Stopwatch.StartNew();
        
        int count = 0;
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            count = worksheet.Dimension?.End.Row ?? 0;
        }
        
        sw.Stop();
        
        Console.WriteLine($"✓ EPPlus Read:              {sw.ElapsedMilliseconds,6} ms (Rows read: {count:N0})");
        return Task.FromResult($"{sw.ElapsedMilliseconds} ms");
    }
    
    // XLS Format Benchmark Methods
    
    private Task<string> BenchmarkWriteNpoiXlsAsync()
    {
        var filePath = Path.Combine(testDataPath, "NPOI_Write.xls");
        var data = GenerateTestData();
        
        var sw = Stopwatch.StartNew();
        try
        {
            using (var workbook = new HSSFWorkbook())
            {
                var sheet = workbook.CreateSheet("Sheet1");
                
                // Write header
                var headerRow = sheet.CreateRow(0);
                var columnNames = data.First().Keys.ToList();
                for (int colIndex = 0; colIndex < columnNames.Count; colIndex++)
                {
                    var cell = headerRow.CreateCell(colIndex);
                    cell.SetCellValue(columnNames[colIndex]);
                }
                
                // Write data
                int rowIndex = 1;
                foreach (var dataRow in data)
                {
                    var row = sheet.CreateRow(rowIndex++);
                    int colIndex = 0;
                    foreach (var kvp in dataRow)
                    {
                        var cell = row.CreateCell(colIndex++);
                        cell.SetCellValue(kvp.Value?.ToString() ?? "");
                    }
                }
                
                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fileStream);
                }
            }
            
            sw.Stop();
            
            var fileSize = new FileInfo(filePath).Length;
            Console.WriteLine($"✓ NPOI XLS Write:           {sw.ElapsedMilliseconds,6} ms (File size: {fileSize / 1024:N0} KB)");
            return Task.FromResult($"{sw.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            sw.Stop();
            Console.WriteLine($"✗ NPOI XLS Write:           Failed - {ex.Message}");
            return Task.FromResult("Failed");
        }
    }
    
    private Task<string> BenchmarkReadMiniExcelXlsAsync()
    {
        // MiniExcel does not support XLS format
        Console.WriteLine($"✗ MiniExcel XLS Read:       Not Supported");
        return Task.FromResult("N/A");
    }
    
    private Task<string> BenchmarkReadExcelDataReaderXlsAsync()
    {
        var filePath = Path.Combine(testDataPath, "NPOI_Write.xls");
        
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"✗ ExcelDataReader XLS Read: Skipped (no XLS file)");
            return Task.FromResult("N/A");
        }
        
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        
        var sw = Stopwatch.StartNew();
        
        try
        {
            int count = 0;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    if (result.Tables.Count > 0)
                    {
                        count = result.Tables[0].Rows.Count;
                    }
                }
            }
            
            sw.Stop();
            
            Console.WriteLine($"✓ ExcelDataReader XLS Read: {sw.ElapsedMilliseconds,6} ms (Rows read: {count:N0})");
            return Task.FromResult($"{sw.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            sw.Stop();
            Console.WriteLine($"✗ ExcelDataReader XLS Read: Failed - {ex.Message}");
            return Task.FromResult("Failed");
        }
    }
    
    private string GetColumnName(int columnNumber)
    {
        string columnName = "";
        
        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }
        
        return columnName;
    }
}
