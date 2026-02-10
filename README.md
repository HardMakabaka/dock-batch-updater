# DOCX Batch Updater

A Windows desktop application for batch updating DOCX documents with strict format preservation.

## Features

- **Batch Processing**: Process up to 100+ DOCX documents simultaneously
- **Format Preservation**: Maintains all formatting attributes including:
  - Font styles (name, size, bold, italic, underline)
  - Text colors and highlights
  - Paragraph alignment and indentation
  - Line spacing
  - Table structures and cell formatting
- **Table Support**: Edit content within tables while preserving structure
- **Auto Backup**: Optional automatic backup of original documents
- **Progress Tracking**: Real-time progress display and detailed logs
- **User-Friendly GUI**: Intuitive interface built with PyQt5
- **Standalone Executable**: Packaged as a single .exe file (no Python required)

## Compatibility

- **Windows 7**
- **Windows 10**
- **Windows 11**

## Installation

### Option 1: Using the Executable (Recommended)

1. Download `DOCX Batch Updater.exe` from the `dist` folder
2. Run the executable directly
3. No additional installation required

### Option 2: Running from Source

#### Prerequisites

- Python 3.8 or higher
- Windows operating system

#### Steps

1. Clone or download the source code
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   cd src
   python main.py
   ```

## Building the Executable

To build the standalone executable:

1. Ensure all dependencies are installed:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the build script (Windows):
   ```bash
   build.bat
   ```

3. The executable will be created in the `dist` folder:
   ```
   dist/DOCX Batch Updater.exe
   ```

## Usage Guide

### Step 1: Add Files

1. Click **Add Files** to select individual DOCX files
2. Or click **Add Folder** to select a folder containing DOCX files
3. The file list will display all selected documents

### Step 2: Define Replacement Rules

1. Enter text to find in the **Find** field
2. Enter replacement text in the **Replace** field
3. Click **Add Rule** to add the replacement rule
4. Repeat for additional rules

**Example:**
- Find: `2024`
- Replace: `2025`
- Result: All occurrences of "2024" will be replaced with "2025"

### Step 3: Configure Backup Options

1. Check **Create Backup Files** to automatically create backups
2. Optional: Click **Select Backup Directory** to specify a backup location
3. If no backup directory is specified, backups will be created in the same folder as the original files

### Step 4: Process Documents

1. Review your files and replacement rules
2. Click **Start Processing**
3. Monitor progress in the progress bar and log
4. Wait for processing to complete

### Step 5: Review Results

After processing:
- The log will show results for each file
- Statistics will display total, successful, and failed counts
- Backups (if enabled) will be created with `_backup` suffix

## Project Structure

```
docx-batch-updater/
├── src/
│   ├── main.py              # Program entry point
│   ├── gui/
│   │   ├── __init__.py
│   │   ├── main_window.py   # Main GUI window
│   │   └── widgets.py       # Custom UI components
│   ├── core/
│   │   ├── __init__.py
│   │   ├── docx_processor.py  # Document processing
│   │   └── batch_processor.py  # Batch processing logic
│   └── utils/
│       ├── __init__.py
│       └── format_preserver.py  # Format preservation utilities
├── tests/                   # Test cases
├── requirements.txt         # Python dependencies
├── build.bat               # Build script
├── README.md               # This file
```

## Technical Details

### Format Preservation Algorithm

The application uses python-docx's run-level operations to preserve formatting:

1. **Text Detection**: Locates text within paragraph runs
2. **Format Capture**: Captures all formatting attributes before replacement
3. **Text Replacement**: Replaces text while preserving run structure
4. **Format Application**: Reapplies captured formatting to the new text

### Batch Processing

- Uses thread pool executor for parallel processing
- Configurable worker threads (default: 4)
- Real-time progress updates
- Comprehensive error handling

### Dependencies

- `python-docx==0.8.11`: DOCX document manipulation
- `PyQt5==5.15.9`: GUI framework
- `PyInstaller==5.13.0`: Executable packaging

## Troubleshooting

### Processing Errors

If processing fails for some documents:

1. Check the log for specific error messages
2. Ensure files are valid DOCX documents
3. Verify files are not corrupted or password-protected
4. Check that files are not open in another application

### Performance Issues

- Large files (>10MB) may take longer to process
- Reduce the number of simultaneous files if experiencing slowdowns
- Close other applications to free system resources

### Build Issues

- Ensure Python 3.8+ is installed
- Verify all dependencies are installed correctly
- Run as administrator if permission issues occur

## Best Practices

1. **Always create backups** when processing important documents
2. **Test with a single file** before batch processing
3. **Verify results** after processing, especially with complex formatting
4. **Keep replacement rules simple** - complex regex patterns may cause issues
5. **Process in batches** if dealing with more than 100 documents

## Limitations

- Only supports `.docx` format (not `.doc`)
- Macros and VBA code are not preserved
- Some advanced Word features (e.g., smart art, embedded objects) may not be fully supported

## Contributing

Contributions are welcome! Please ensure:

1. Code follows PEP 8 style guidelines
2. All functions have proper docstrings
3. Tests are included for new features
4. Changes are documented in the commit message

## License

This project is provided as-is for educational and commercial use.

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Review the log output for specific error details
3. Ensure your documents meet the requirements

## Changelog

### Version 1.0.0
- Initial release
- Batch file processing
- Format preservation
- Table support
- Backup functionality
- Progress tracking
- Standalone executable packaging
