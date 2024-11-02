<?php

include_once 'excel.class.php';

$excel = new Excel();

// Set filename
$excel->file('my-file.xlsx');

// Write first sheet
$contactsData = [
   ['Name', 'Age', 'Email'],
   ['John Doe', '31', 'john@example.com'],
   ['Jane Smith', '25', 'jane@example.com']
];

$excel->write('Contacts', $contactsData);

// Write second sheet
$productsData = [
   ['Product', 'Price', 'Quantity'],
   ['Laptop', '1000', '5'],
   ['Smartphone', '500', '10']
];

$excel->write('Products', $productsData);

// Get list of all sheets
$sheetNames = $excel->sheets();

// Read all data from all sheets
$allData = $excel->read();

// Read specific sheet by name
$contactsByName = $excel->read('Contacts');

// Read specific sheet by index (0-based)
$contactsByIndex = $excel->read(0);

// Reading using first row as headers
$contactsWithHeaders = $excel->read('Contacts', true);
$productsWithHeaders = $excel->read('Products', true);