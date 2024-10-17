# Directory Tree Excel Generator

A Node.js script that generates an Excel file containing a structured tree view of files and directories with status tracking capabilities. The script supports file/directory exclusions and extension filtering while maintaining directory structures.

## Installation

1. Create a new directory and initialize a Node.js project:

```bash
mkdir directory-tree-excel
cd directory-tree-excel
npm init -y
```

2. Install required dependencies:

```bash
npm install exceljs yargs
```

3. Save the script as `index.js` in your project directory.

## Usage

### Basic Usage

```bash
node index.js --path "/path/to/your/directory"
```

### With Extension Filtering

```bash
node index.js --path "/path/to/your/directory" --include-only-extensions "js,jsx,ts,tsx"
```

## Command Line Arguments

| Argument | Alias | Description | Required | Default |
|----------|-------|-------------|----------|---------|
| `--path` | `-p` | Root directory path to generate tree from | Yes | - |
| `--include-only-extensions` | `-e` | Comma-separated list of file extensions to include | No | "" (all files) |

## Exclusion List

The script looks for a file named `TO_EXCLUDE_TREE_TEMP.txt` in the root of the provided path. This file should contain a list of files and directories to exclude from the tree.

### Example `TO_EXCLUDE_TREE_TEMP.txt`:

```bash
file.txt
node_modules/
.git/
dist/
build/
```

Notes:

- Use forward slashes (`/`) for directory paths
- Add a trailing slash (`/`) to exclude directories
- Paths should be relative to the root directory
- The `TO_EXCLUDE_TREE_TEMP.txt` file itself is automatically excluded

## Output Excel Structure

The script generates a file named `directory-tree.xlsx` in the current working directory with the following columns:

| Column | Description |
|--------|-------------|
| Level | Numeric value indicating the depth in the directory tree (0 = root) |
| Type | Either "File" or "Directory" |
| Name | Item name with proper indentation showing hierarchy |
| Path | Relative path from the root directory |
| Status | Dropdown with options: pending, processing, done |

### Features:

- Auto-filtering enabled on all columns
- Status column includes data validation (dropdown)
- Directory names are in bold
- Proper indentation showing hierarchy
- Relative paths for better portability
- Directories end with a trailing slash

## Examples

1. Generate tree for a project including only TypeScript files:

```bash
node index.js -p "./my-project" -e "ts,tsx"
```

2. Generate complete tree excluding node_modules:

```bash
# Create exclusion file
echo "node_modules/" > TO_EXCLUDE_TREE_TEMP.txt
# Run script
node index.js -p "./my-project"
```

## Notes

- The script works across all platforms (Windows, macOS, Linux)
- Paths in the Excel file use forward slashes for consistency
- Directories without any matching files (when using extension filtering) are excluded
- The directory structure is preserved even when filtering by extension
- All paths in the output are relative to the provided root path

## Error Handling

The script includes error handling for:

- Invalid paths
- Missing permissions
- Malformed exclusion files
- File system errors

Errors are logged to the console with appropriate context.
