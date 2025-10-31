# ExcelReadAndWriteRace

A comprehensive performance comparison tool for popular .NET Excel libraries.

## Overview

This project benchmarks the read and write performance of five major Excel libraries for .NET:

- **MiniExcel 1.41.4** - A lightweight, fast Excel file processing library
- **ExcelDataReader 3.8.0** - A read-only library for reading Excel files
- **DocumentFormat.OpenXml 3.3.0** - Microsoft's official Open XML SDK
- **ClosedXML 0.105.0** - A high-level Excel library built on top of OpenXML
- **EPPlus 4.5.3.3** - A popular library for creating Excel files using .NET

## Features

- Write performance benchmarks for 10,000 rows × 10 columns
- Read performance benchmarks for the same dataset
- File size comparison
- Formatted output with performance summary table

## Requirements

- .NET 8.0 SDK or later

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
WRITE PERFORMANCE TESTS
===============================================================================

✓ MiniExcel Write:             428 ms (File size: 715 KB)
✓ OpenXml Write:               896 ms (File size: 572 KB)
✓ ClosedXML Write:            1582 ms (File size: 666 KB)
✓ EPPlus Write:               1046 ms (File size: 653 KB)

===============================================================================
READ PERFORMANCE TESTS
===============================================================================

✓ MiniExcel Read:              379 ms (Rows read: 10,001)
✓ ExcelDataReader Read:        717 ms (Rows read: 10,001)
✓ OpenXml Read:                800 ms (Rows read: 10,001)
✓ ClosedXML Read:             2348 ms (Rows read: 10,001)
✓ EPPlus Read:                 593 ms (Rows read: 10,001)

===============================================================================
SUMMARY
===============================================================================

┌─────────────────────────────────────┬─────────────────┬─────────────────┐
│ Library                             │ Write Time      │ Read Time       │
├─────────────────────────────────────┼─────────────────┼─────────────────┤
│ MiniExcel 1.41.4                    │ 428 ms          │ 379 ms          │
│ ExcelDataReader 3.8.0               │ N/A (Read-only) │ 717 ms          │
│ DocumentFormat.OpenXml 3.3.0        │ 896 ms          │ 800 ms          │
│ ClosedXML 0.105.0                   │ 1582 ms         │ 2348 ms         │
│ EPPlus 4.5.3.3                      │ 1046 ms         │ 593 ms          │
└─────────────────────────────────────┴─────────────────┴─────────────────┘
```

## Key Findings

### Write Performance
1. **MiniExcel** - Fastest write performance (428 ms)
2. **DocumentFormat.OpenXml** - 2.1x slower than MiniExcel
3. **EPPlus** - 2.4x slower than MiniExcel
4. **ClosedXML** - 3.7x slower than MiniExcel
5. **ExcelDataReader** - Read-only library, no write support

### Read Performance
1. **MiniExcel** - Fastest read performance (379 ms)
2. **EPPlus** - 1.6x slower than MiniExcel
3. **ExcelDataReader** - 1.9x slower than MiniExcel
4. **DocumentFormat.OpenXml** - 2.1x slower than MiniExcel
5. **ClosedXML** - 6.2x slower than MiniExcel

### File Size
- **DocumentFormat.OpenXml** - Smallest file size (572 KB)
- **EPPlus** - 653 KB
- **ClosedXML** - 666 KB
- **MiniExcel** - Largest file size (715 KB)

## Libraries Used

| Library | Version | Purpose | License |
|---------|---------|---------|---------|
| MiniExcel | 1.41.4 | Fast read/write | MIT |
| ExcelDataReader | 3.8.0 | Read-only | MIT |
| DocumentFormat.OpenXml | 3.3.0 | Microsoft's OpenXML SDK | MIT |
| ClosedXML | 0.105.0 | High-level read/write | MIT |
| EPPlus | 4.5.3.3 | Popular read/write library | LGPL |

## Conclusion

**MiniExcel** emerges as the clear winner for both read and write performance, making it an excellent choice for applications that prioritize speed. **EPPlus** provides a good balance between performance and features, with competitive read speeds. However, the choice of library should also consider:

- **Feature requirements** - ClosedXML and EPPlus offer rich feature sets for complex Excel operations
- **File size** - DocumentFormat.OpenXml produces the smallest files
- **Simplicity** - ExcelDataReader is excellent for read-only scenarios
- **Official support** - DocumentFormat.OpenXml is Microsoft's official library
- **Licensing** - EPPlus uses LGPL license (v4.5.3.3), while others are MIT licensed

## License

MIT License
