<?php

namespace Core\Misc;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Fill;

use Exception;

class XlsxUpdater
{
    private $sheets = [];
    private $writer;
    private $pathToFileXlsx;
    private $spreadSheet;

    public function __construct(string $pathToFileXlsx) {
        try {
            $this->pathToFileXlsx = $pathToFileXlsx;
            $this->spreadSheet = IOFactory::load($pathToFileXlsx);
            $this->writer = IOFactory::createWriter($this->spreadSheet, 'Xlsx');
            $this->sheets = $this->spreadSheet->getAllSheets();
        } catch (Exception $e) {
            print_r($e->getMessage());
        }
    }

    public function newBooking(string $hash, string $nameGuest, string $dateStart, string $dateEnd, string $colorRGB = 'F28A8C') {
        try {
            $arrayOfDates = $this->genArrayOfDates($dateStart,$dateEnd);
            $arrayOfSheetsNames = $this->genArrayOfSheetsNames($dateStart,$dateEnd);
            $arrayOfSheetsForRecord = $this->getSheetsForRecord($arrayOfSheetsNames);
            $idRoom = $this->getIdRoom($hash);
            if ($idRoom == null || $arrayOfDates == null) {
                throw new Exception();
            }
            foreach ($arrayOfSheetsForRecord as $sheet) {
                foreach ($arrayOfDates as $date) {
                    $month = trim($sheet->getTitle());
                    if (stripos($date,$month)) {
                        $cells = $sheet->getCellCollection();
                        $day = date('j',strtotime($date)) + 2;
                        $cell = trim($cells->get($idRoom.$day)->getValue());
                        if ($cell == '' || is_numeric($cell)) {
                            echo $cell;
                            $cells->get($idRoom.$day)->setValue($nameGuest);
                            $cells->get($idRoom.$day)->getStyle()
                                                     ->getFill()
                                                     ->setFillType(Fill::FILL_SOLID)
                                                     ->getStartColor()
                                                     ->setRGB($colorRGB);
                        } else {
                            throw new Exception();
                        }
                    }
                }
            }
            $this->writer->save($this->pathToFileXlsx);
            return true;
        } catch (Exception $e) {
            return false;
        }

    }
    
    public function getIdRoom(string $hash)
    {
        for ($letter = 'B'; $letter <= 'K'; $letter++) {
            if (md5($letter) == $hash) {
                return $letter;
            }
        }
        return null;
    }
    
    public function getSheetsDocument() {
        $array = [];
        foreach ($this->sheets as $sheet) {
            $array[] = $sheet->getTitle();
        }
        return $array;
    }

    private function genArrayOfDates(string $dateStart, string $dateEnd)
    {
        $array = [];
        $dateStart = strtotime($dateStart);
        $dateEnd = strtotime($dateEnd);
        if (($dateEnd - $dateStart) < 0) {
            return null;
        }
        for ($i = $dateStart; $i <= $dateEnd; $i += 86400) {
            $array[] = date("d.m.Y", $i);
        }
        return $array;
    }

    private function genArrayOfSheetsNames(string $dateStart, string $dateEnd)
    {
        $array = [];
        $dateStart = strtotime($dateStart);
        $dateEnd = strtotime($dateEnd);
        for ($i = $dateStart; $i <= $dateEnd; $i += 86400) {
            $date = date("m.Y", $i);
            if (!in_array($date, $array)) {
                $array[] = $date;
            }
        }
        return $array;
    }

    private function getSheetsForRecord(Array $arrayOfSheets)
    {
        $array = [];
        foreach ($this->sheets as $sheet) {
            if (in_array(trim($sheet->getTitle()),$arrayOfSheets)) {
                $array[] = $sheet;
            }
        }
        return $array;
    }
}

?>
