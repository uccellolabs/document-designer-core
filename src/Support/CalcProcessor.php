<?php

namespace Uccello\DocumentDesignerCore\Support;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class CalcProcessor
{
    protected $spreadsheet;
    protected $sheetNames;
    protected $variablesIndex;

    public function __construct($templateFile)
    {
        $reader = IOFactory::createReader('Xlsx');
        //$reader->setReadDataOnly(TRUE);

        $this->spreadsheet = $reader->load($templateFile);
        $this->sheetNames = $this->spreadsheet->getSheetNames();

        $this->initVariablesIndex();
    }

    public function process($data)
    {
        $data = $this->sortData($data);

        foreach ($data as $key => $value) {
            $this->processNode($key, $value);
        }
    }

    public function saveAs($outFile)
    {
        $writer = IOFactory::createWriter($this->spreadsheet, 'Xlsx');
        $writer->save($outFile);
    }

    public function getVariables(array $options = null)
    {
        $this->initVariablesIndex();

        if (is_array($options) && in_array('sortBySheet', $options)) { // Option: sortBySheet
            $variables = [];

            foreach ($this->variablesIndex as $varName => $varData) {
                foreach ($varData as $sheet => $cell) {
                    $variables[$this->sheetNames[$sheet]][] = $varName;
                }
            }
        } else { // No option
            $variables = array_keys($this->variablesIndex);
        }

        return $variables;
    }

    private function initVariablesIndex()
    {
        $this->variablesIndex = [];

        foreach ($this->sheetNames as $sheetIndex => $sheetName) {
            // dump("SHEET: $sheetName");
            $worksheet = $this->spreadsheet->getSheet($sheetIndex);

            foreach ($worksheet->getRowIterator() as $row) {
                $cellIterator = $row->getCellIterator();
                // $cellIterator->setIterateOnlyExistingCells(FALSE);  // This loops through all cells,
                //    even if a cell value is not set.
                // By default, only cells that have a value
                //    set will be iterated.
                foreach ($cellIterator as $cell) {
                    $val = $cell->getValue();
                    $coords = $cell->getCoordinate();

                    if (!empty($val)) {
                        $variables = $this->getVariablesForCell($val);

                        foreach ($variables as $variable) {
                            // dump("[$sheetIndex][$coords] $val => $variable");
                            $this->variablesIndex[$variable][$sheetIndex][] = $coords;
                        }
                    }
                }
            }
        }

        // dump($this->variablesIndex);
    }

    protected function getVariablesForCell($cellValue)
    {
        preg_match_all('/\$\{(.*?)(:[^}]*)*\}/i', $cellValue, $matches);

        return $matches[1];
    }

    private function processNode($key, $data)
    {
        $type   = substr($key, 0, 2);

        if ($type == 'b:' || $type == 's:' || $type == 't:' || $type == 'i:') {
            $key = substr($key, 2);
        }

        if ($type == 'b:') { // Blocks
            // Ignore: Only for Docx Templates
            return false;
        } elseif ($type == 's:') { // Sheet
            // TODO !!!
        } elseif ($type == 't:') { // Table
            $this->processRow($key, $data);
        } elseif ($type == 'i:') { // Image
            $this->setImageValue($key, $data);
        } else { // Vars
            $this->setValue($key, $data);
        }

        return true;
    }

    protected function setValue($search, $replace) // TODO : Add limit param...
    {
        if (!empty($this->variablesIndex[$search]) && is_array($this->variablesIndex[$search])) {
            foreach ($this->variablesIndex[$search] as $sheetIndex => $cells) {
                // dump("SHEET: $sheetIndex");
                $worksheet = $this->spreadsheet->getSheet($sheetIndex);

                foreach ($cells as $coords) {
                    $cell = $worksheet->getCell($coords);

                    $this->setValueForCell($cell, $search, $replace);
                }
            }
        }
    }

    private function setValueForCell(&$cell, $search, $replace)
    {
        $val = $cell->getValue();

        $val = preg_replace('/\$\{' . $search . '(:[^}]*)*\}/i', $replace, $val); // TODO: Limit..

        // TODO : Add limit param:
        // // Note: we can't use the same function for both cases here, because of performance considerations.
        // if (self::MAXIMUM_REPLACEMENTS_DEFAULT === $limit) {
        //     return str_replace($search, $replace, $documentPartXML);
        // }
        // $regExpEscaper = new RegExp();

        // return preg_replace($regExpEscaper->escape($search), $replace, $documentPartXML, $limit);

        // dump("[$search][".$cell->getCoordinate()."] $val");

        $cell->setValue($val);
    }

    private function processRow($key, $data)
    {
        if (isset($this->variablesIndex[$key]) && is_array($this->variablesIndex[$key])) {
            foreach ($this->variablesIndex[$key] as $sheetIndex => $cells) {
                // dump("SHEET: $sheetIndex");
                $worksheet = $this->spreadsheet->getSheet($sheetIndex);

                foreach ($cells as $coords) {
                    $tempRowIndex = Coordinate::coordinateFromString($coords)[1];

                    $worksheet->insertNewRowBefore($tempRowIndex + 1, count($data));

                    $highestColumn = $worksheet->getHighestColumn();
                    $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);

                    for ($columnIndex = 1; $columnIndex <= $highestColumnIndex; ++$columnIndex) {
                        $tempValue = $worksheet->getCellByColumnAndRow($columnIndex, $tempRowIndex)->getValue();

                        if (!empty($tempValue)) {
                            foreach ($data as $i => $vars) {
                                $cell = $worksheet->getCellByColumnAndRow($columnIndex, $tempRowIndex + $i + 1);
                                $cell->setValue($tempValue);

                                foreach ($vars as $search => $replace) {
                                    $type   = substr($search, 0, 2);

                                    if ($type == 'i:') { // Image
                                        $search = substr($search, 2);

                                        // TODO: Handle !!!
                                    } else { // Vars
                                        $this->setValueForCell($cell, $search, $replace);
                                    }
                                }
                            }
                        }
                    }

                    $worksheet->removeRow($tempRowIndex);
                }
            }
        }
    }

    protected function sortData($data)
    {
        $sheets = [];
        $tables = [];
        $images = [];
        $vars   = [];

        foreach ($data as $key => $value) {
            $type = substr($key, 0, 2);

            if ($type == 's:') {
                $sheets[$key] = $value;
            }
            if ($type == 't:') {
                $tables[$key] = $value;
            }
            if ($type == 'i:') {
                $images[$key] = $value;
            } else {
                $vars[$key] = $value;
            }
        }

        return array_merge($sheets, $tables, $images, $vars);
    }

    protected function setImageValue($search, $replace) // TODO : Add limit param...
    {
        if (is_array($this->variablesIndex[$search])) {
            foreach ($this->variablesIndex[$search] as $sheetIndex => $cells) {
                // dump("SHEET: $sheetIndex");
                $worksheet = $this->spreadsheet->getSheet($sheetIndex);

                foreach ($cells as $coords) {
                    $cell = $worksheet->getCell($coords);

                    $this->setImageValueForCell($cell, $search, $replace);
                }
            }
        }
    }

    protected function setImageValueForCell(&$cell, $search, $replace)
    {
        $val = $cell->getValue();

        $matches = null;

        preg_match('/\$\{' . $search . '(:[^}]*)*\}/i', $val, $matches);

        $fullVariable   = $matches[0];
        $args           = $matches[1];

        $imageArgs = $this->getImageArgs($args);

        $drawing = new Drawing();
        // $drawing->setName('Image');
        // $drawing->setDescription('Image');
        $drawing->setPath($replace); // put your path and image here
        $drawing->setCoordinates($cell->getCoordinate());
        $drawing->setWidth($imageArgs['width']);
        $drawing->setHeight($imageArgs['height']);
        // $drawing->setOffsetX(110);
        // $drawing->setRotation(25);
        // $drawing->getShadow()->setVisible(true);
        // $drawing->getShadow()->setDirection(45);
        $drawing->setWorksheet($this->spreadsheet->getActiveSheet());


        $val = str_replace($fullVariable, '', $val); // TODO: Limit..

        // dump("[$search][" . $cell->getCoordinate() . "] $val");

        $cell->setValue($val);
    }

    // Code from PHPWord
    protected function getImageArgs($varNameWithArgs)
    {
        $varElements = explode(':', $varNameWithArgs);
        array_shift($varElements); // first element is name of variable => remove it

        $varInlineArgs = array();
        // size format documentation: https://msdn.microsoft.com/en-us/library/documentformat.openxml.vml.shape%28v=office.14%29.aspx?f=255&MSPPError=-2147217396
        foreach ($varElements as $argIdx => $varArg) {
            if (strpos($varArg, '=')) { // arg=value
                list($argName, $argValue) = explode('=', $varArg, 2);
                $argName = strtolower($argName);
                if ($argName == 'size') {
                    list($varInlineArgs['width'], $varInlineArgs['height']) = explode('x', $argValue, 2);
                } else {
                    $varInlineArgs[strtolower($argName)] = $argValue;
                }
            } elseif (preg_match('/^([0-9]*[a-z%]{0,2}|auto)x([0-9]*[a-z%]{0,2}|auto)$/i', $varArg)) { // 60x40
                list($varInlineArgs['width'], $varInlineArgs['height']) = explode('x', $varArg, 2);
            } else { // :60:40:f
                switch ($argIdx) {
                    case 0:
                        $varInlineArgs['width'] = $varArg;
                        break;
                    case 1:
                        $varInlineArgs['height'] = $varArg;
                        break;
                    case 2:
                        $varInlineArgs['ratio'] = $varArg;
                        break;
                }
            }
        }

        return $varInlineArgs;
    }
}
