/*const express = require('express');
const XLSX = require('xlsx');
const chokidar = require('chokidar');
const path = require('path');

const app = express();
const PORT = 3000;

let excelData = [];

// Function to load Excel data
const loadExcelData = () => {
    const filePath = path.join(__dirname, 'data.xlsx');
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    excelData = XLSX.utils.sheet_to_json(worksheet);
};

// Function to perform structured search in Excel data
function structuredSearch(query) {
    // Parse query into key, value, and condition
    const parts = query.split(',');
    if (parts.length !== 3) {
        throw new Error('Invalid query format. Expected format: key,value,condition');
    }
    
    const [key, valueStr, condition] = parts.map(part => part.trim());
    let value = valueStr;

    // Convert value to number if possible
    if (!isNaN(value)) {
        value = parseFloat(value);
    }

    // Perform search based on key, value, and condition
    const searchResults = excelData.filter(row => {
        // Case insensitive comparison for 'name'
        if (key.toLowerCase() === 'name') {
            const rowValue = row[key].toString().toLowerCase();
            const searchValue = value.toString().toLowerCase();

            switch (condition) {
                case 'equals':
                    return rowValue === searchValue;
                case 'not equals':
                    return rowValue !== searchValue;
                default:
                    throw new Error(`Unsupported condition '${condition}' for key '${key}'`);
            }
        }

        // Numeric comparison for other keys
        const rowValue = parseFloat(row[key]);

        switch (condition) {
            case 'equals':
                return rowValue === value;
            case 'not equals':
                return rowValue !== value;
            case 'greater than':
                return rowValue > value;
            case 'less than':
                return rowValue < value;
            case 'greater than or equals':
                return rowValue >= value;
            case 'less than or equals':
                return rowValue <= value;
            default:
                throw new Error(`Unsupported condition '${condition}' for key '${key}'`);
        }
    });

    return searchResults;
}

// Serve static files from the 'public' directory
app.use(express.static(path.join(__dirname, 'public')));

// Endpoint to load all data
app.get('/alldata', (req, res) => {
    res.json(excelData);
});

// Endpoint to handle structured search requests
app.get('/search', (req, res) => {
    const { query } = req.query;
    if (!query) {
        return res.status(400).json({ error: 'Query parameter is required' });
    }

    try {
        // Perform structured search
        const results = structuredSearch(query);
        res.json(results);
    } catch (error) {
        console.error('Error performing search:', error.message);
        res.status(400).json({ error: error.message });
    }
});

// Load Excel data initially
loadExcelData();


// Start server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
*/

const express = require('express');
const XLSX = require('xlsx');
const chokidar = require('chokidar');
const path = require('path');

const app = express();
const PORT = 3000;
const dataFilePath = path.join(__dirname, 'data.xlsx'); // Define data file path

let excelData = [];

// Function to load Excel data
const loadExcelData = () => {
    const workbook = XLSX.readFile(dataFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    excelData = XLSX.utils.sheet_to_json(worksheet);
};

// Function to perform structured search in Excel data
function structuredSearch(query) {
    // Parse query into key, value, and condition
    const parts = query.split(',');
    if (parts.length !== 3) {
        throw new Error('Invalid query format. Expected format: key,value,condition');
    }
    
    const [key, valueStr, condition] = parts.map(part => part.trim());
    let value = valueStr;

    // Convert value to number if possible
    if (!isNaN(value)) {
        value = parseFloat(value);
    }

    // Perform search based on key, value, and condition
    const searchResults = excelData.filter(row => {
        // Case insensitive comparison for 'name'
        if (key.toLowerCase() === 'name') {
            const rowValue = row[key].toString().toLowerCase();
            const searchValue = value.toString().toLowerCase();

            switch (condition) {
                case 'equals':
                    return rowValue === searchValue;
                case 'not equals':
                    return rowValue !== searchValue;
                default:
                    throw new Error(`Unsupported condition '${condition}' for key '${key}'`);
            }
        }

        // Numeric comparison for other keys
        const rowValue = parseFloat(row[key]);

        switch (condition) {
            case 'equals':
                return rowValue === value;
            case 'not equals':
                return rowValue !== value;
            case 'greater than':
                return rowValue > value;
            case 'less than':
                return rowValue < value;
            case 'greater than or equals':
                return rowValue >= value;
            case 'less than or equals':
                return rowValue <= value;
            default:
                throw new Error(`Unsupported condition '${condition}' for key '${key}'`);
        }
    });

    return searchResults;
}

// Serve static files from the 'public' directory
app.use(express.static(path.join(__dirname, 'public')));

// Endpoint to load all data
app.get('/alldata', (req, res) => {
    res.json(excelData);
});

// Endpoint to handle structured search requests
app.get('/search', (req, res) => {
    const { query } = req.query;
    if (!query) {
        return res.status(400).json({ error: 'Query parameter is required' });
    }

    try {
        // Perform structured search
        const results = structuredSearch(query);
        res.json(results);
    } catch (error) {
        console.error('Error performing search:', error.message);
        res.status(400).json({ error: error.message });
    }
});

// Load Excel data initially
loadExcelData();

// Watch for changes in data.xlsx
const watcher = chokidar.watch(dataFilePath, {
    persistent: true,
    awaitWriteFinish: {
        stabilityThreshold: 2000,
        pollInterval: 100
    }
});

// Reload data on file change
watcher.on('change', (filePath) => {
    console.log(`File ${filePath} has been changed`);
    loadExcelData();
});

// Start server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
