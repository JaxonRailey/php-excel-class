# PHP Excel Class
A simple PHP class for reading and writing Excel (XLSX) files.

## Requirements
- PHP 8.0 or higher
- ZIP Extension
- SimpleXML Extension

## Usage

### Download and include the class in your project

```php
include_once 'excel.class.php';
```

### Initializalize the istance

```php
$excel = new Excel();
```

### Set or get Excel filename

```php
// Set filename
$excel->file('my-file.xlsx');

// Get current filename
$filename = $excel->file();
```

### Writing data

```php
// Example data
$data = [
    ['Name', 'Age', 'Email'],
    ['John Doe', '31', 'john@example.com'],
    ['Jane Smith', '25', 'jane@example.com']
];

// Write data to a sheet
$excel->write('Contacts', $data);
```

### Get list of all sheets

```php
// Get list of all sheets
$sheetNames = $excel->sheets();
```

### Reading Data

```php
// Read all data from all sheets
$allData = $excel->read();

// Read specific sheet by name
$data = $excel->read('Contacts');

//  Read specific sheet by index (0-based)
$data = $excel->read('Contacts', 1);

// Reading using first row as headers
$data = $excel->read('Contacts', true);
```

The ```read()``` method offers flexible ways to retrieve data from your Excel file:

When called without parameters, it returns all sheets with their data organized by sheet names.

You can specify a particular sheet using either its name or index (0-based).

The second parameter ```useFirstRowAsKeys``` allows you to use the first row values as array keys.

---

:star: **If you liked what I did, if it was useful to you or if it served as a starting point for something more magical let me know with a star** :green_heart:
