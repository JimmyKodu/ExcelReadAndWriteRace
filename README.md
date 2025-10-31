# ExcelReadAndWriteRace

A comprehensive performance comparison tool for popular .NET Excel libraries.

## Overview

This project benchmarks the read and write performance of five major Excel libraries for .NET:

- **MiniExcel 1.41.4** - A lightweight, fast Excel file processing library
- **ExcelDataReader 3.8.0** - A read-only library for reading Excel files
- **DocumentFormat.OpenXml 3.3.0** - Microsoft's official Open XML SDK
- **ClosedXML 0.105.0** - A high-level Excel library built on top of OpenXML
- **EPPlus 4.5.3.3** - A popular library for creating Excel files using .NET
- **NPOI 2.7.3** - A .NET port of Apache POI for XLS/XLSX file support

The benchmark includes tests for both modern XLSX format and legacy XLS (Excel 97-2003) format to evaluate compatibility.

## Features

- Write performance benchmarks for 10,000 rows × 10 columns
- Read performance benchmarks for the same dataset
- File size comparison
- Formatted output with performance summary table
- XLSX format (modern Excel) support testing
- XLS format (Excel 97-2003) compatibility testing

## Requirements

- .NET 9.0 SDK or later

## Installation

```bash
git clone https://github.com/JimmyKodu/ExcelReadAndWriteRace.git
cd ExcelReadAndWriteRace/ExcelReadAndWriteRace
dotnet build
```

## Usage

```bash
dotnet run
```

## Benchmark Results

Sample output from running the benchmark:

```
================================================================================
Excel Libraries Performance Comparison
================================================================================

Test Configuration:
  Rows: 10,000
  Columns: 10

===============================================================================
WRITE PERFORMANCE TESTS (XLSX Format)
===============================================================================

✓ MiniExcel Write:             301 ms (File size: 1,093 KB)
✓ OpenXml Write:               771 ms (File size: 573 KB)
✓ ClosedXML Write:            1483 ms (File size: 663 KB)
✓ EPPlus Write:                849 ms (File size: 653 KB)

===============================================================================
READ PERFORMANCE TESTS (XLSX Format)
===============================================================================

✓ MiniExcel Read:              294 ms (Rows read: 10,001)
✓ ExcelDataReader Read:        565 ms (Rows read: 10,001)
✓ OpenXml Read:                803 ms (Rows read: 10,001)
✓ ClosedXML Read:             1190 ms (Rows read: 10,001)
✓ EPPlus Read:                 551 ms (Rows read: 10,001)

===============================================================================
XLS FORMAT (Excel 97-2003) COMPATIBILITY TESTS
===============================================================================

✓ NPOI XLS Write:              501 ms (File size: 2,737 KB)
✓ ExcelDataReader XLS Read:    410 ms (Rows read: 10,001)
✗ MiniExcel XLS Read:       Not Supported

Note: Only ExcelDataReader and NPOI support the legacy XLS format.
      DocumentFormat.OpenXml, ClosedXML, EPPlus, and MiniExcel only support XLSX.

===============================================================================
SUMMARY - XLSX Format
===============================================================================

┌─────────────────────────────────────┬─────────────────┬─────────────────┐
│ Library                             │ Write Time      │ Read Time       │
├─────────────────────────────────────┼─────────────────┼─────────────────┤
│ MiniExcel 1.41.4                    │ 301 ms          │ 294 ms          │
│ ExcelDataReader 3.8.0               │ N/A (Read-only) │ 565 ms          │
│ DocumentFormat.OpenXml 3.3.0        │ 771 ms          │ 803 ms          │
│ ClosedXML 0.105.0                   │ 1483 ms         │ 1190 ms         │
│ EPPlus 4.5.3.3                      │ 849 ms          │ 551 ms          │
└─────────────────────────────────────┴─────────────────┴─────────────────┘

===============================================================================
SUMMARY - XLS Format (Excel 97-2003)
===============================================================================

┌─────────────────────────────────────┬─────────────────┬─────────────────┐
│ Library                             │ Write Time      │ Read Time       │
├─────────────────────────────────────┼─────────────────┼─────────────────┤
│ MiniExcel 1.41.4                    │ N/A             │ N/A             │
│ ExcelDataReader 3.8.0               │ N/A (Read-only) │ 410 ms          │
│ DocumentFormat.OpenXml 3.3.0        │ N/A             │ N/A             │
│ ClosedXML 0.105.0                   │ N/A             │ N/A             │
│ EPPlus 4.5.3.3                      │ N/A             │ N/A             │
│ NPOI 2.7.3                          │ 501 ms          │ N/A             │
└─────────────────────────────────────┴─────────────────┴─────────────────┘
```

## Key Findings

### XLSX Format Performance

#### Write Performance
1. **MiniExcel** - Fastest write performance (301 ms)
2. **DocumentFormat.OpenXml** - 2.6x slower than MiniExcel
3. **EPPlus** - 2.8x slower than MiniExcel
4. **ClosedXML** - 4.9x slower than MiniExcel
5. **ExcelDataReader** - Read-only library, no write support

#### Read Performance
1. **MiniExcel** - Fastest read performance (294 ms)
2. **EPPlus** - 1.9x slower than MiniExcel
3. **ExcelDataReader** - 1.9x slower than MiniExcel
4. **DocumentFormat.OpenXml** - 2.7x slower than MiniExcel
5. **ClosedXML** - 4.0x slower than MiniExcel

#### File Size
- **DocumentFormat.OpenXml** - Smallest file size (573 KB)
- **EPPlus** - 653 KB
- **ClosedXML** - 663 KB
- **MiniExcel** - Largest file size (1,093 KB)

### XLS Format (Excel 97-2003) Compatibility

**Key Finding:** Most modern libraries do NOT support the legacy XLS binary format.

#### Supported Libraries:
- **NPOI** - Full read/write support for XLS files (Write: 501 ms, File size: 2,737 KB)
- **ExcelDataReader** - Read-only support for XLS files (Read: 410 ms)

#### Not Supported:
- **MiniExcel** - No XLS support
- **DocumentFormat.OpenXml** - XLSX only
- **ClosedXML** - XLSX only
- **EPPlus** - XLSX only

**Recommendation:** For legacy XLS file support, use NPOI (read/write) or ExcelDataReader (read-only).

## Libraries Used

| Library | Version | Purpose | Format Support | License |
|---------|---------|---------|----------------|---------|
| MiniExcel | 1.41.4 | Fast read/write | XLSX, CSV | MIT |
| ExcelDataReader | 3.8.0 | Read-only | XLSX, XLS, CSV | MIT |
| DocumentFormat.OpenXml | 3.3.0 | Microsoft's OpenXML SDK | XLSX only | MIT |
| ClosedXML | 0.105.0 | High-level read/write | XLSX only | MIT |
| EPPlus | 4.5.3.3 | Popular read/write library | XLSX only | LGPL |
| NPOI | 2.7.3 | Apache POI port for .NET | XLSX, XLS | Apache 2.0 |

## Conclusion

### For XLSX (Modern Excel) Files:
**MiniExcel** emerges as the clear winner for both read and write performance, making it an excellent choice for applications that prioritize speed. **EPPlus** provides a good balance between performance and features, with competitive read speeds.

### For XLS (Legacy Excel 97-2003) Files:
**Only NPOI and ExcelDataReader** support the legacy XLS format. If you need to work with old XLS files:
- Use **NPOI** for read/write operations
- Use **ExcelDataReader** for read-only operations (faster than NPOI for reading)

### Selection Criteria:

The choice of library should consider:

- **Format compatibility** - Need XLS support? Use NPOI or ExcelDataReader
- **Performance** - MiniExcel for XLSX, ExcelDataReader for XLS reading
- **Feature requirements** - ClosedXML and EPPlus offer rich feature sets for complex Excel operations
- **File size** - DocumentFormat.OpenXml produces the smallest XLSX files
- **Simplicity** - ExcelDataReader is excellent for read-only scenarios
- **Official support** - DocumentFormat.OpenXml is Microsoft's official library
- **Licensing** - EPPlus uses LGPL license (v4.5.3.3), NPOI uses Apache 2.0, others are MIT licensed

## License

MIT License
