<?php

class Excel {

    protected array $sheets = [];
    protected array $sharedStrings = [];
    protected string $filename;


    /**
     * Sets or gets current file
     *
     * @param string $filename - path to the Excel file (optional)
     *
     * @return mixed returns Excel instance or the current filename
     */

    public function file(string $filename = null) {

        if ($filename !== null) {
            $this->filename = $filename;
            return $this;
        }

        return $this->filename;
    }


    /**
     * Writes sheet data
     *
     * @param string $sheet - name of the sheet
     * @param array $data - data to write
     *
     * @return bool
     */

    public function write(string $sheet, array $data) {

        if (!$this->filename) {
            throw new Exception('File not set. Use file() method to set a file first');
        }

        $this->addSheet($sheet, $data);
        $this->save();

        return true;
    }


    /**
     * Reads sheet data
     *
     * @param mixed $sheet - name or id of the sheet (optional)
     * @param bool $useFirstRowAsKeys - if keys are attribute of first row (optional)
     *
     * @return array returns sheet data or all sheets data
     */

    public function read(mixed $sheet = null, bool $useFirstRowAsKeys = false) {

        if (!$this->filename) {
            throw new Exception('File not set. Use file() method to set a file first');
        }

        $zip = new ZipArchive();

        if ($zip->open($this->filename) !== true) {
            throw new Exception('Unable to open ' . $this->filename);
        }

        $data = [];
        $sheetMap = $this->getSheetMap($zip);
        $this->loadSharedStrings($zip);

        $sheetsToRead = [];
        if ($sheet === null) {
            $sheetsToRead = $sheetMap;
        } elseif (is_int($sheet)) {
            $sheetsToRead = [$sheetMap[$sheet] ?? null];
        } elseif (is_string($sheet)) {
            $sheetsToRead = array_filter($sheetMap, fn($s) => $s['name'] === $sheet);
        } else {
            throw new Exception('$sheet parameter must be null, an integer, or a string');
        }

        foreach ($sheetsToRead as $sheetInfo) {
            if ($sheetInfo === null) continue;
            $sheetData = $this->readSheetData($zip, $sheetInfo['path'], $useFirstRowAsKeys);
            $data[$sheetInfo['name']] = $sheetData;
        }

        $zip->close();

        return $data;
    }


    /**
     * Returns the list of sheet names in Excel file
     *
     * @return array array with sheet names
     */

    public function sheets() {

        if (!$this->filename) {
            throw new Exception('File not set. Use file() method to set a file first');
        }

        $zip = new ZipArchive();

        if ($zip->open($this->filename) !== true) {
            throw new Exception('Unable to open ' . $this->filename);
        }

        $sheetNames  = [];
        $workbookXml = $zip->getFromName('xl/workbook.xml');

        if (!$workbookXml) {
            throw new Exception("Unable to read 'xl/workbook.xml' from Excel file");
        }

        $xml = new SimpleXMLElement($workbookXml);
        $xml->registerXPathNamespace('ns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $sheets = $xml->xpath('//ns:sheets/ns:sheet');

        foreach ($sheets as $sheet) {
            $attributes = $sheet->attributes();
            $sheetNames[] = (string)$attributes['name'];
        }

        $zip->close();

        return $sheetNames;
    }


    // Private Methods

    protected function save() {

        if (!$this->filename) {
            throw new Exception('File not set. Use file() method to set a file first');
        }

        $files = $this->generateFiles();
        $zip   = new ZipArchive();

        if ($zip->open($this->filename, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
            throw new Exception('Unable to open ' . $this->filename);
        }

        foreach ($files as $file => $xml) {
            $zip->addFromString($file, $xml);
        }

        $zip->close();
    }

    private function addSheet(string $sheet, array $data) {

        $this->sheets[] = ['name' => $sheet, 'data' => $data];
    }

    private function loadSharedStrings(ZipArchive $zip) {

        $xml = $zip->getFromName('xl/sharedStrings.xml');

        if ($xml === false) {
            return;
        }

        $xml = new SimpleXMLElement($xml);
        $xml->registerXPathNamespace('ns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        foreach ($xml->xpath('//ns:si') as $string) {
            $this->sharedStrings[] = (string)$string->t;
        }
    }

    private function generateFiles() {

        $files = [
            '[Content_Types].xml' => $this->generateContentTypesXML(),
            '_rels/.rels' => $this->generateRelsXML(),
            'xl/workbook.xml' => $this->generateWorkbookXML(),
            'xl/_rels/workbook.xml.rels' => $this->generateWorkbookRelsXML(),
            'xl/styles.xml' => $this->generateStylesXML(),
        ];

        foreach ($this->sheets as $index => $sheet) {
            $sheetXML = $this->generateSheetXML($sheet['data']);
            $files['xl/worksheets/sheet' . ($index + 1) . '.xml'] = $sheetXML;
        }

        return $files;
    }

    private function generateContentTypesXML() {

        $sheetsOverrides = '';

        foreach ($this->sheets as $index => $sheet) {
            $sheetsOverrides .= '<Override PartName="/xl/worksheets/sheet' . ($index + 1) . '.xml" ' . 'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }

        $xml = '<?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                ' . $sheetsOverrides . '
                <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
            </Types>';

        return $xml;
    }

    private function generateRelsXML() {

        $xml = '<?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>';

        return $xml;
    }

    private function generateWorkbookXML() {

        $sheetsXML = '';

        foreach ($this->sheets as $index => $sheet) {
            $sheetName = htmlspecialchars($sheet['name'], ENT_QUOTES | ENT_XML1);
            $sheetsXML .= '<sheet name="' . $sheetName . '" sheetId="' . ($index + 1) . '" r:id="rId' . ($index + 1) . '"/>';
        }

        $xml = '<?xml version="1.0" encoding="UTF-8"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets>' . $sheetsXML . '</sheets>
            </workbook>';

        return $xml;
    }

    private function generateWorkbookRelsXML() {

        $relationships = '';

        foreach ($this->sheets as $index => $sheet) {
            $relationships .= '<Relationship Id="rId' . ($index + 1) . '" ' .
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ' .
                'Target="worksheets/sheet' . ($index + 1) . '.xml"/>';
        }

        $styleRelId = 'rId' . (count($this->sheets) + 1);

        $relationships .= '<Relationship Id="' . $styleRelId . '" ' .
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" ' .
            'Target="styles.xml"/>';

        $xml = '<?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                ' . $relationships . '
            </Relationships>';

        return $xml;
    }

    private function generateStylesXML() {

        $xml = '<?xml version="1.0" encoding="UTF-8"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <fonts count="1">
                    <font>
                        <sz val="11"/>
                        <color theme="1"/>
                        <name val="Calibri"/>
                        <family val="2"/>
                    </font>
                </fonts>
                <fills count="1">
                    <fill>
                        <patternFill patternType="none"/>
                    </fill>
                </fills>
                <borders count="1">
                    <border>
                        <left/><right/><top/><bottom/><diagonal/>
                    </border>
                </borders>
                <cellStyleXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                </cellStyleXfs>
                <cellXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                </cellXfs>
                <cellStyles count="1">
                    <cellStyle name="Normal" xfId="0" builtinId="0"/>
                </cellStyles>
                <dxfs count="0"/>
                <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
            </styleSheet>';

        return $xml;
    }

    private function generateSheetXML(array $data) {

        $rowsXML = '';

        foreach ($data as $rowIndex => $row) {
            $cellsXML = '';
            foreach ($row as $colIndex => $cellValue) {
                $cellRef = $this->num2cell($colIndex) . ($rowIndex + 1);
                $cellValueEscaped = htmlspecialchars((string)$cellValue, ENT_QUOTES | ENT_XML1);
                $cellsXML .= '<c r="' . $cellRef . '" t="inlineStr"><is><t>' . $cellValueEscaped . '</t></is></c>';
            }
            $rowsXML .= '<row r="' . ($rowIndex + 1) . '">' . $cellsXML . '</row>';
        }

        $xml = '<?xml version="1.0" encoding="UTF-8"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>' . $rowsXML . '</sheetData>
            </worksheet>';

        return $xml;
    }

    private function getSheetMap(ZipArchive $zip) {

        $workbookXml = $zip->getFromName('xl/workbook.xml');

        if ($workbookXml === false) {
            throw new Exception("Unable to read 'xl/workbook.xml' from Excel file");
        }

        $xml = new SimpleXMLElement($workbookXml);
        $xml->registerXPathNamespace('ns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $xml->registerXPathNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $sheets = $xml->xpath('//ns:sheets/ns:sheet');

        $sheetMap = [];

        foreach ($sheets as $sheet) {
            $sheetId    = (string)$sheet['sheetId'];
            $sheetName  = (string)$sheet['name'];
            $rId        = (string)$sheet->attributes('r', true)['id'];
            $sheetRel   = $this->getSheetRelation($zip, $rId);
            $sheetPath  = 'xl/' . $sheetRel;

            $sheetMap[] = [
                'id'   => $sheetId,
                'name' => $sheetName,
                'path' => $sheetPath,
            ];
        }

        return $sheetMap;
    }

    private function getSheetRelation(ZipArchive $zip, string $rId){

        $workbookRelsXml = $zip->getFromName('xl/_rels/workbook.xml.rels');

        if ($workbookRelsXml === false) {
            throw new Exception("Unable to read 'xl/_rels/workbook.xml.rels' from Excel file");
        }

        $xml = new SimpleXMLElement($workbookRelsXml);
        $xml->registerXPathNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships');
        $relationship = $xml->xpath('//rel:Relationship[@Id="' . $rId . '"]');

        if (empty($relationship)) {
            throw new Exception("Relation with Id '" . $rId . "' not found");
        }

        return (string)$relationship[0]['Target'];
    }

    private function readSheetData(ZipArchive $zip, string $sheetPath, bool $useFirstRowAsKeys) {

        $sheetXml = $zip->getFromName($sheetPath);

        if ($sheetXml === false) {
            throw new Exception("Unable to read '" . $sheetPath . "' from Excel file");
        }

        $xml = new SimpleXMLElement($sheetXml);
        $xml->registerXPathNamespace('ns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        $rows   = [];
        $header = [];

        foreach ($xml->sheetData->row as $row) {
            $rowData = [];
            foreach ($row->c as $cell) {
                $cellValue = $this->getCellValue($cell);
                $colRef = preg_replace('/\d/', '', (string)$cell['r']);
                $colIndex = $this->cell2num($colRef);
                $rowData[$colIndex] = $cellValue;
            }

            ksort($rowData);
            $rowData = array_values($rowData);

            if ($useFirstRowAsKeys && empty($header)) {
                $header = $rowData;
            } elseif ($useFirstRowAsKeys) {
                $rows[] = array_combine($header, $rowData);
            } else {
                $rows[] = $rowData;
            }
        }

        return $rows;
    }

    private function getCellValue(SimpleXMLElement $cell) {

        $type = (string)$cell['t'];

        if ($type === 's') {
            $index = (int)$cell->v;
            return $this->sharedStrings[$index] ?? '';
        } elseif ($type === 'inlineStr') {
            return (string)$cell->is->t;
        } else {
            return (string)$cell->v;
        }
    }

    private function num2cell(int $num) {

        $letters = '';

        while ($num >= 0) {
            $letters = chr($num % 26 + 65) . $letters;
            $num = intdiv($num, 26) - 1;
        }

        return $letters;
    }

    private function cell2num(string $cellRef) {

        $num = 0;
        $len = strlen($cellRef);

        for ($i = 0; $i < $len; $i++) {
            $num *= 26;
            $num += ord($cellRef[$i]) - 64;
        }

        return $num - 1;
    }
}