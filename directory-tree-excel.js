const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');

// Parse command line arguments
const argv = yargs(hideBin(process.argv))
    .option('path', {
        alias: 'p',
        type: 'string',
        description: 'Path to generate tree from',
        demandOption: true
    })
    .option('include-only-extensions', {
        alias: 'e',
        type: 'string',
        description: 'Comma-separated list of file extensions to include (e.g., "js,jsx,ts,tsx")',
        default: ''
    })
    .argv;

// Parse included extensions
const includedExtensions = argv['include-only-extensions']
    ? new Set(argv['include-only-extensions'].split(',').map(ext => ext.toLowerCase().trim()))
    : null;

// Initialize Excel workbook and worksheet
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Directory Tree');

// Set up columns
worksheet.columns = [
    { header: 'Level', key: 'level', width: 10 },
    { header: 'Type', key: 'type', width: 10 },
    { header: 'Name', key: 'name', width: 50 },
    { header: 'Path', key: 'path', width: 100 },
    { header: 'Status', key: 'status', width: 15 }
];

// Add styles for the header
worksheet.getRow(1).font = { bold: true };
worksheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E0E0' }
};

// Status dropdown list
const statusValidation = {
    type: 'list',
    allowBlank: false,
    formulae: ['"pending,processing,done"']
};

let rowIndex = 2;

// Function to normalize path for current platform
function normalizePath(inputPath) {
    return inputPath.split(/[/\\]/).filter(Boolean).join(path.sep);
}

// Function to read exclusion list
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

// Function to check if path should be excluded
function shouldExclude(itemPath, excludeList) {
    const normalizedItemPath = normalizePath(itemPath);
    
    // Check direct match
    if (excludeList.has(normalizedItemPath)) {
        return true;
    }

    // Check if path is within an excluded directory
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

// Function to check if file should be included based on extension
function shouldIncludeFile(filePath) {
    // If no extension filter is set, include all files
    if (!includedExtensions) {
        return true;
    }

    const ext = path.extname(filePath).toLowerCase().slice(1);
    return includedExtensions.has(ext);
}

// Function to check if directory contains any files with included extensions
function hasIncludedFiles(dirPath) {
    // If no extension filter is set, include all directories
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

// Function to get relative path
function getRelativePath(fullPath, basePath) {
    const relativePath = path.relative(basePath, fullPath);
    return relativePath.split(path.sep).join('/');
}

// Function to process directory
function processDirectory(dirPath, basePath, excludeList, level = 0) {
    try {
        const items = fs.readdirSync(dirPath);
        const dirHasIncludedFiles = hasIncludedFiles(dirPath);

        // Skip directory if it doesn't contain any included files
        if (!dirHasIncludedFiles) {
            return;
        }

        for (const item of items) {
            const fullPath = path.join(dirPath, item);
            const relativePath = getRelativePath(fullPath, basePath);

            // Skip if item should be excluded
            if (shouldExclude(relativePath, excludeList)) {
                continue;
            }

            const stats = fs.statSync(fullPath);
            const isDirectory = stats.isDirectory();

            // Skip files that don't match the extension filter
            if (!isDirectory && !shouldIncludeFile(item)) {
                continue;
            }

            // Skip directories that don't contain any included files
            if (isDirectory && !hasIncludedFiles(fullPath)) {
                continue;
            }

            // Add row to worksheet
            const row = worksheet.getRow(rowIndex);
            row.getCell('level').value = level;
            row.getCell('type').value = isDirectory ? 'Directory' : 'File';
            row.getCell('name').value = '  '.repeat(level) + item;
            row.getCell('path').value = relativePath + (isDirectory ? '/' : '');
            row.getCell('status').value = 'pending';

            // Add data validation for status column
            worksheet.getCell(`E${rowIndex}`).dataValidation = statusValidation;

            // Style the row
            row.font = { size: 11 };
            if (isDirectory) {
                row.font.bold = true;
            }

            rowIndex++;

            // If it's a directory, process its contents
            if (isDirectory) {
                processDirectory(fullPath, basePath, excludeList, level + 1);
            }
        }
    } catch (error) {
        console.error(`Error processing directory ${dirPath}:`, error);
    }
}

async function generateExcel() {
    try {
        const basePath = path.resolve(argv.path);
        const excludeList = getExclusionList(basePath);

        console.log('Excluded items:', Array.from(excludeList));
        if (includedExtensions) {
            console.log('Including only files with extensions:', Array.from(includedExtensions));
        }

        // Process the directory
        processDirectory(basePath, basePath, excludeList);

        // Auto-filter for all columns
        worksheet.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: rowIndex - 1, column: 5 }
        };

        // Generate the Excel file
        const outputPath = path.join(process.cwd(), 'directory-tree.xlsx');
        await workbook.xlsx.writeFile(outputPath);
        console.log(`Excel file generated successfully at: ${outputPath}`);
    } catch (error) {
        console.error('Error generating Excel file:', error);
    }
}

// Run the script
generateExcel();
