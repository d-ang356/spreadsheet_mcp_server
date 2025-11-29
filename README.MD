# Spreadsheet MCP Server

A comprehensive Model Context Protocol (MCP) server for managing Excel and CSV spreadsheets with advanced features including formulas, formatting, charts, and pane freezing.

## Features

### Core Operations
- **File Management**: Create, read, write, delete, rename, and list spreadsheet files
- **Sheet Operations**: Create, rename, and manage multiple sheets within workbooks
- **Data Manipulation**: Read, write, append rows, and update individual cells
- **Formula Support**: Set and retrieve Excel formulas with cross-sheet references

### Advanced Features
- **Cell Formatting**: Bold, italic, font size, and background colors
- **Range Formatting**: Format multiple cells at once
- **Column/Row Formatting**: Set column widths and row heights
- **Freeze Panes**: Freeze rows and/or columns for easier data navigation
- **Charts**: Create bar and pie charts from data ranges

### Supported Formats
- **Excel (.xlsx)**: Full feature support including formulas, formatting, and charts
- **CSV (.csv)**: Basic read/write operations

## Installation

### Prerequisites
- Docker installed and running
- Claude Desktop or MCP-compatible client

### Setup

1. **Clone or create the project directory**:
```bash
mkdir spreadsheet-mcp-server
cd spreadsheet-mcp-server
```

2. **Create the required files**:
   - `spreadsheet_server.py` (main server code)
   - `requirements.txt` (Python dependencies)
   - `Dockerfile` (container configuration)
   - `catalog.yaml` (MCP catalog configuration)

3. **Install Python dependencies** (for local testing):
```bash
pip install -r requirements.txt
```

4. **Build the Docker image**:
```bash
docker build -t spreadsheet-mcp-server:latest .
```

5. **Configure Claude Desktop**:

Add to your MCP settings:
```json
{
  "mcpServers": {
    "spreadsheet": {
      "command": "docker",
      "args": [
        "run",
        "-i",
        "--rm",
        "--init",
        "--stop-signal=SIGTERM",
        "-v",
        "/path/to/your/spreadsheets:/app/spreadsheets",
        "-v",
        "/path/to/imports:/imports",
        "spreadsheet-mcp-server:latest"
      ]
    }
  }
}
```

Replace `/path/to/your/spreadsheets` with your desired working directory.

## Usage Examples

### Creating a Spreadsheet
```python
# Create a new Excel file with headers
await create_spreadsheet(
    filename="sales.xlsx",
    sheet_name="Q1",
    headers=["Date", "Product", "Amount"]
)
```

### Writing Data
```python
# Write data to a specific sheet (creates sheet if it doesn't exist)
await write_spreadsheet(
    filename="sales.xlsx",
    sheet="Q1",
    data=[
        ["Date", "Product", "Amount"],
        ["2024-01-01", "Widget", 100],
        ["2024-01-02", "Gadget", 150]
    ]
)
```

### Using Formulas
```python
# Set a formula in a cell
await set_formula(
    filename="sales.xlsx",
    sheet="Summary",
    cell="B2",
    formula="=Q1!C2"
)

# Cross-sheet reference
await set_formula(
    filename="sales.xlsx",
    sheet="Data1",
    cell="E2",
    formula="=Data2!A2"
)
```

### Formatting
```python
# Format header row with green background
await format_cells(
    filename="sales.xlsx",
    sheet="Q1",
    cell_range="A1:E1",
    bold=True,
    bg_color="#00FF00"
)

# Format a single cell
await set_cell_format(
    filename="sales.xlsx",
    sheet="Q1",
    cell="A1",
    bold=True,
    italic=True,
    bg_color="FFFF0000"  # Red with full opacity
)
```

### Column and Row Sizing
```python
# Set column width
await set_column_format(
    filename="sales.xlsx",
    sheet="Q1",
    column="A",
    width=20
)

# Set row height
await set_row_format(
    filename="sales.xlsx",
    sheet="Q1",
    row=1,
    height=30
)
```

### Freezing Panes
```python
# Freeze top row
await freeze_panes(
    filename="sales.xlsx",
    sheet="Q1",
    cell="A2"
)

# Freeze first column
await freeze_panes(
    filename="sales.xlsx",
    sheet="Q1",
    cell="B1"
)

# Freeze both top row and first column
await freeze_panes(
    filename="sales.xlsx",
    sheet="Q1",
    cell="B2"
)

# Unfreeze panes
await unfreeze_panes(
    filename="sales.xlsx",
    sheet="Q1"
)
```

### Creating Charts
```python
# Create a bar chart
await create_chart(
    filename="sales.xlsx",
    sheet="Q1",
    chart_type="bar",
    data_range="A1:B10",
    title="Sales by Product"
)

# Create a pie chart
await create_chart(
    filename="sales.xlsx",
    sheet="Q1",
    chart_type="pie",
    data_range="A1:B5",
    title="Market Share"
)
```

## API Reference

### File Operations

| Tool | Description | Required Parameters |
|------|-------------|-------------------|
| `list_files` | List all spreadsheet files | `pattern` (optional) |
| `create_spreadsheet` | Create new spreadsheet | `filename` |
| `delete_spreadsheet` | Delete a spreadsheet | `filename` |
| `rename_file` | Rename a file | `old_filename`, `new_filename` |

### Sheet Operations

| Tool | Description | Required Parameters |
|------|-------------|-------------------|
| `rename_sheet` | Rename a sheet | `filename`, `old_sheet`, `new_sheet` |

### Data Operations

| Tool | Description | Required Parameters |
|------|-------------|-------------------|
| `read_spreadsheet` | Read spreadsheet data | `filename` |
| `write_spreadsheet` | Write data to spreadsheet | `filename`, `data` |
| `append_row` | Append a single row | `filename`, `row_data` |
| `update_cell` | Update single cell | `filename`, `sheet`, `cell`, `value` |

### Formula Operations

| Tool | Description | Required Parameters |
|------|-------------|-------------------|
| `set_formula` | Set cell formula | `filename`, `sheet`, `cell`, `formula` |
| `get_formula` | Get cell formula | `filename`, `sheet`, `cell` |

### Formatting Operations

| Tool | Description | Required Parameters |
|------|-------------|-------------------|
| `format_cells` | Format cell range | `filename`, `sheet`, `cell_range` |
| `set_cell_format` | Format single cell | `filename`, `sheet`, `cell` |
| `set_column_format` | Format column | `filename`, `sheet`, `column` |
| `set_row_format` | Format row | `filename`, `sheet`, `row` |

### Advanced Features

| Tool | Description | Required Parameters |
|------|-------------|-------------------|
| `freeze_panes` | Freeze panes | `filename`, `sheet`, `cell` |
| `unfreeze_panes` | Remove frozen panes | `filename`, `sheet` |
| `create_chart` | Create chart | `filename`, `sheet`, `chart_type`, `data_range` |

## Color Formatting

Colors can be specified in two formats:

1. **RGB Hex** (6 characters): `#00FF00` or `00FF00`
   - Automatically converted to ARGB with full opacity
   
2. **ARGB Hex** (8 characters): `FF00FF00`
   - First two characters = alpha (opacity): `FF` = fully opaque
   - Remaining six = RGB color

Examples:
- Green: `#00FF00` → converted to `FF00FF00`
- Red: `#FF0000` → converted to `FFFF0000`
- Semi-transparent blue: `8000FF00` (50% opacity)

## Directory Structure

```
spreadsheet-mcp-server/
├── spreadsheet_server.py    # Main server code
├── requirements.txt          # Python dependencies
├── Dockerfile               # Docker configuration
├── catalog.yaml            # MCP catalog
├── README.md               # This file
└── spreadsheets/           # Working directory (created automatically)
```

## Troubleshooting

### Server not appearing in Claude Desktop
1. Verify Docker image is built: `docker images | grep spreadsheet-mcp-server`
2. Check Claude Desktop logs: `~/Library/Logs/Claude/mcp*.log` (Mac) or `%APPDATA%\Claude\logs\mcp*.log` (Windows)
3. Ensure all tools have `inputSchema` defined in `handle_tools_list()`

### "Colors must be aRGB hex values" error
- Use the `_normalize_color()` helper function
- Ensure colors are in `#RRGGBB` or `AARRGGBB` format

### "Worksheet does not exist" error
- The `write_spreadsheet` and `append_row` methods now auto-create sheets
- Rebuild Docker image after updating code

### Sheet not created
- Ensure you're using the updated `write_spreadsheet()` that includes auto-sheet creation
- Verify the sheet name doesn't contain invalid characters

## Development

### Running Tests Locally
```bash
# Install dependencies
pip install -r requirements.txt

# Run server in stdio mode
python spreadsheet_server.py
```

### Rebuilding After Changes
```bash
# Rebuild Docker image
docker build -t spreadsheet-mcp-server:latest .

# Restart Claude Desktop completely (quit and reopen)
```

## Version History

### v1.1.0 (Current)
- Added freeze/unfreeze panes functionality
- Auto-create sheets in write operations
- Color normalization (RGB → ARGB)
- Chart creation (bar and pie)
- Column and row formatting
- Enhanced error handling

### v1.0.0
- Initial release
- Basic CRUD operations
- Formula support
- Cell formatting

## License

MIT License - feel free to modify and distribute.

## Contributing

Contributions are welcome! Please ensure:
1. All tools have proper `inputSchema` definitions
2. Error handling is implemented
3. Docker image builds successfully
4. Tools are tested with Claude Desktop

## Support

For issues or questions:
- Check the troubleshooting section
- Review Claude Desktop logs
- Verify Docker container is running: `docker ps | grep spreadsheet`