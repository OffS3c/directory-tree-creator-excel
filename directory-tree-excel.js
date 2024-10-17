/**
 * Directory Tree Excel Generator
 * 
 * This script generates an Excel file containing a structured tree view of files and directories
 * with status tracking capabilities. It supports file/directory exclusions and extension filtering
 * while maintaining directory structures.
 * 
 * @author OffS3c
 * @version 1.0.0
 */

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');

// Parse command line arguments with detailed help messages
const argv = yargs(hideBin(process.argv))
    .option('path', {
        alias: 'p',
        type: 'string',
        description: 'Root directory path to generate tree from',
        demandOption: true
    })
    .option('include-only-extensions', {
        alias: 'e',
        type: 'string',
        description: 'Comma-separated list of file extensions to include (e.g., "js,jsx,ts,tsx")',
        default: ''
    })
    .example('$0 --path "./my-project"', 'Generate tree for all files')
    .example('$0 --path "./my-project" -e "js,jsx,ts,tsx"', 'Generate tree for JavaScript/TypeScript files only')
    .argv;

/**
 * Set of file extensions to include in the tree.
 * If null, all files will be included.
 * @type {Set<string>|null}
 */
const includedExtensions = argv['include-only-extensions']
    ? new Set(argv['include-only-extensions'].split(',').map(ext => ext.toLowerCase().trim()))
    : null;

// Initialize Excel workbook and worksheet configuration
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Directory Tree');

// Configure worksheet columns with proper width and headers
worksheet.columns = [
    { header: 'Level', key: 'level', width: 10 },
    { header: 'Type', key: 'type', width: 10 },
    { header: 'Name', key: 'name', width: 50 },
    { header: 'Path', key: 'path', width: 100 },
    { header: 'Status', key: 'status', width: 15 }
];

// Style header row
worksheet.getRow(1).font = { bold: true };
worksheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E0E0' }
};

// Configure status column validation
const statusValidation = {
    type: 'list',
    allowBlank: false,
    formulae: ['"pending,processing,done"']
};

let rowIndex = 2;

/**
 * Normalizes a path string to use the current platform's path separator
 * @param {string} inputPath - The path to normalize
 * @returns {string} Normalized path using platform-specific separators
 */
function normalizePath(inputPath) {
    return inputPath.split(/[/\\]/).filter(Boolean).join(path.sep);
}

/**
 * Reads and parses the exclusion list file
 * @param {string} basePath - The root directory path
 * @returns {Set<string>} Set of paths to exclude
 */
function getExclusionList(basePath) {
    const excludeFile = path.join(basePath, 'TO_EXCLUDE_TREE_TEMP.txt');
    const excludeList = new Set(['TO_EXCLUDE_TREE_TEMP.txt']);

    try {
        if (fs.existsSync(excludeFile)) {
            const content = fs.readFileSync(excludeFile, 'utf8');
            content.split('\n').forEach(line => {
                const trimmedLine = line.trim();
                if (trimmedLine) {
                    excludeList.add(normalizePath(trimmedLine));
                }
            });
        }
    } catch (error) {
        console.error('Error reading exclusion file:', error);
    }

    return excludeList;
}

/**
 * Checks if a path should be excluded based on the exclusion list
 * @param {string} itemPath - Path to check
 * @param {Set<string>} excludeList - Set of paths to exclude
 * @returns {boolean} True if the path should be excluded
 */
function shouldExclude(itemPath, excludeList) {
    const normalizedItemPath = normalizePath(itemPath);
    
    if (excludeList.has(normalizedItemPath)) {
        return true;
    }

    for (const excludePath of excludeList) {
        if (excludePath.endsWith(path.sep)) {
            const excludeDir = excludePath.slice(0, -1);
            if (normalizedItemPath === excludeDir || 
                normalizedItemPath.startsWith(excludeDir + path.sep)) {
                return true;
            }
        }
    }

    return false;
}

/**
 * Checks if a file should be included based on its extension
 * @param {string} filePath - Path of the file to check
 * @returns {boolean} True if the file should be included
 */
function shouldIncludeFile(filePath) {
    if (!includedExtensions) {
        return true;
    }

    const ext = path.extname(filePath).toLowerCase().slice(1);
    return includedExtensions.has(ext);
}

/**
 * Recursively checks if a directory contains any files with included extensions
 * @param {string} dirPath - Directory path to check
 * @returns {boolean} True if the directory contains matching files
 */
function hasIncludedFiles(dirPath) {
    if (!includedExtensions) {
        return true;
    }

    let hasMatchingFiles = false;
    
    function checkDir(currentPath) {
        const items = fs.readdirSync(currentPath);
        
        for (const item of items) {
            const fullPath = path.join(currentPath, item);
            const stats = fs.statSync(fullPath);
            
            if (stats.isDirectory()) {
                if (checkDir(fullPath)) {
                    hasMatchingFiles = true;
                    return true;
                }
            } else if (shouldIncludeFile(item)) {
                hasMatchingFiles = true;
                return true;
            }
        }
        
        return hasMatchingFiles;
    }

    return checkDir(dirPath);
}

/**
 * Converts an absolute path to a path relative to the base directory
 * @param {string} fullPath - Absolute path to convert
 * @param {string} basePath - Base directory path
 * @returns {string} Relative path using forward slashes
 */
function getRelativePath(fullPath, basePath) {
    const relativePath = path.relative(basePath, fullPath);
    return relativePath.split(path.sep).join('/');
}

/**
 * Recursively processes a directory and adds its contents to the Excel worksheet
 * @param {string} dirPath - Current directory path
 * @param {string} basePath - Base directory path
 * @param {Set<string>} excludeList - Set of paths to exclude
 * @param {number} level - Current directory depth
 */
function processDirectory(dirPath, basePath, excludeList, level = 0) {
    try {
        const items = fs.readdirSync(dirPath);
        const dirHasIncludedFiles = hasIncludedFiles(dirPath);

        if (!dirHasIncludedFiles) {
            return;
        }

        for (const item of items) {
            const fullPath = path.join(dirPath, item);
            const relativePath = getRelativePath(fullPath, basePath);

            if (shouldExclude(relativePath, excludeList)) {
                continue;
            }

            const stats = fs.statSync(fullPath);
            const isDirectory = stats.isDirectory();

            if (!isDirectory && !shouldIncludeFile(item)) {
                continue;
            }

            if (isDirectory && !hasIncludedFiles(fullPath)) {
                continue;
            }

            // Add item to worksheet with proper formatting
            const row = worksheet.getRow(rowIndex);
            row.getCell('level').value = level;
            row.getCell('type').value = isDirectory ? 'Directory' : 'File';
            row.getCell('name').value = '  '.repeat(level) + item;
            row.getCell('path').value = relativePath + (isDirectory ? '/' : '');
            row.getCell('status').value = 'pending';

            worksheet.getCell(`E${rowIndex}`).dataValidation = statusValidation;

            row.font = { size: 11 };
            if (isDirectory) {
                row.font.bold = true;
            }

            rowIndex++;

            if (isDirectory) {
                processDirectory(fullPath, basePath, excludeList, level + 1);
            }
        }
    } catch (error) {
        console.error(`Error processing directory ${dirPath}:`, error);
    }
}

/**
 * Main function to generate the Excel file
 * Initializes the process and handles the file generation
 */
async function generateExcel() {
    try {
        const basePath = path.resolve(argv.path);
        const excludeList = getExclusionList(basePath);

        console.log('Excluded items:', Array.from(excludeList));
        if (includedExtensions) {
            console.log('Including only files with extensions:', Array.from(includedExtensions));
        }

        processDirectory(basePath, basePath, excludeList);

        worksheet.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: rowIndex - 1, column: 5 }
        };

        const outputPath = path.join(process.cwd(), 'directory-tree.xlsx');
        await workbook.xlsx.writeFile(outputPath);
        console.log(`Excel file generated successfully at: ${outputPath}`);
    } catch (error) {
        console.error('Error generating Excel file:', error);
    }
}

// Execute the script
generateExcel();
