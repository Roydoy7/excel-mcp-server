# Excel MCP Server - Gemini CLI Extension

This is an integration of [excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) as a Gemini CLI built-in extension.

## Features

The Excel MCP Server provides comprehensive Excel file manipulation capabilities:

- **Workbook Operations**: Create, read, update workbooks and worksheets
- **Data Manipulation**: Write and read data, formulas, formatting
- **Charts**: Create various chart types (line, bar, pie, scatter, etc.)
- **Pivot Tables**: Create dynamic pivot tables for data analysis
- **Formatting**: Font styling, colors, borders, alignment, conditional formatting
- **Tables**: Create and manage Excel tables with custom styling
- **Sheet Management**: Copy, rename, delete worksheets

## How It Works

This extension integrates the Python-based Excel MCP Server into Gemini CLI:

1. **Automatic Installation**: When you run `npm install`, the `postinstall` script automatically installs all Python dependencies using the embedded Python 3.13.7
2. **Zero Configuration**: The extension is automatically loaded by Gemini CLI's extension system
3. **MCP Server**: Runs as an MCP server using stdio transport for local file manipulation

## File Structure

```
excel-mcp-server/
├── src/
│   └── excel_mcp/           # Original Python source code
├── gemini-extension.json    # Extension configuration
├── package.json            # NPM package configuration
├── install-deps.js         # Automatic dependency installer
└── EXTENSION.md           # This file
```

## Usage

Once installed, you can use Excel tools through Gemini CLI:

### Example Commands (via LLM)

- "Create a new Excel file at C:/temp/test.xlsx"
- "Write this data to the first sheet: [data]"
- "Format cells A1 to C3 with bold text and blue background"
- "Create a bar chart from the data in A1:B10"
- "Add a pivot table to analyze sales by region"

## Available Tools

See [TOOLS.md](TOOLS.md) for complete documentation of all 40+ available tools.

## Dependencies

Python dependencies (automatically installed):
- mcp[cli] >= 1.10.1
- fastmcp >= 2.0.0, < 3.0.0
- openpyxl >= 3.1.5
- typer >= 0.16.0

## Platform Support

- ✅ Windows (win32)
- ✅ macOS (darwin)
- ✅ Linux

## License

This extension integrates the original excel-mcp-server which is licensed under MIT License.
