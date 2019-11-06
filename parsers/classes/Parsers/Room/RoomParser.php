<?php

namespace Parsers\Room;

use Websm\Framework\Mail\HTMLMessage;
use Websm\Framework\Chpu;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

use Model\Catalog\Product;

use Exceptions\FileNotFoundException;
use Exceptions\InvalidFileException;
use Exception;

use Rs\Json\Patch;
use Rs\Json\Patch\InvalidPatchDocumentJsonException;
use Rs\Json\Patch\InvalidTargetDocumentJsonException;
use Rs\Json\Patch\InvalidOperationException;

class RoomParser
{
    private const ALLOWED_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    private $xlsxCells;
    private $xlsxSheet;
    private $pathToProductXlsx;
    private $wrongColorRGB = ['000000','FFFFFF'];
    private $actualSheets = [];
    private $sheets;

    public function __construct(string $pathToProductXlsx)
    {
        try {
            $this->checkXlsxFile($pathToProductXlsx);
            $this->pathToProductXlsx = $pathToProductXlsx;
            $reader = new Xlsx();
            $spreadSheet = $reader->load($pathToProductXlsx);
            $this->sheets = $spreadSheet->getAllSheets();
            $this->xlsxSheet = $spreadSheet->getActiveSheet();
            $this->xlsxCells = $this->xlsxSheet->getCellCollection();
            $flag = false;

            foreach ($this->sheets as $sheet) {
                if (date("m.Y") == $sheet->getTitle()) {
                    $flag = true;
                }
                if ($flag) {
                    array_push($this->actualSheets, $sheet->getTitle());
                }
            }
        } catch (FileNotFoundException $e) {
            self::sendMessageError();
            print_r($e->getMessage());
        } catch (InvalidFileException $e) {
            self::sendMessageError();
            print_r($e->getMessage());
        } catch (\Exception $e) {
            self::sendMessageError();
            print_r($e->getMessage());
        }
    }

    static private function sendMessageError()
    {
        $message = new HTMLMessage;

        $message->setFrom('<noreply@mail.websm.io>')
                ->setTo('<developer@websm.io>')
                ->setSubject('Ошибка в работе парсера')
                ->setBody('<h1></h1>');

        $message->send();
    }

    private function checkXlsxFile(string $path)
    {
        if (!file_exists($path)) {
            throw new FileNotFoundException(
                'File on path "'.$path.'" not found' 
            );
        }

        $fileMimeType = mime_content_type($path);
        if ($fileMimeType !== self::ALLOWED_MIME_TYPE) {
            throw new InvalidFileException(
                'Invalid file mime type'
            );
        }
    }

    private function getNameSheet()
    {
        return $this->xlsxSheet->getTitle();
    }

    private function getColorCell(string $cell)
    {
        return $this->xlsxSheet->getStyle($cell)
            ->getFill()
            ->getStartColor()
            ->getRGB();
    }

    private function getCountDaysInSheet()
    {
        $col = 'A';
        $row = 3;
        $count = 0;
        while (!in_array($this->getColorCell($col.$row), $this->wrongColorRGB)) {
            $count++;
            $row++;
        }
        return $count;
    }

    private function getValueCell(string $cell)
    {
        return trim($this->xlsxCells->get($cell)
            ->getValue());
    }
    
    private function getDataRooms()
    {
        $rooms = [];
        for ($col = 'B'; $col <= 'K'; $col++) {
            $roomName = $this->getValueCell($col.'2');
            $room = [];
            foreach ($this->sheets as $sheet) {
                if (in_array($sheet->getTitle(),$this->actualSheets)) {
                    $this->xlsxCells = $sheet->getCellCollection();
                    $this->xlsxSheet = $sheet;
                    for ($row = 3; $row <= $this->getCountDaysInSheet()+2; $row++) {
                        if (!in_array($this->getColorCell($col.$row),$this->wrongColorRGB)) {
                            if ($this->getValueCell($col.$row)) {
                                $guestName = $this->getValueCell($col.$row);
                            }

                            $date = $this->getValueCell('A'.$row).'.'.$this->getNameSheet();
                            array_push($room, [
                                'name' => $guestName,
                                'date' => $date,
                            ]);
                        }
                    }
                }
            }
            $room = [
                'id' => md5($col),
                'title' => $roomName,
                'dates' => $room,
            ];
            array_push($rooms,$room);
        }
        return $rooms;
    }

    public function parse()
    {
        try {
            $rooms = $this->getDataRooms();
            foreach ($rooms as $room) {
                $title = $room['title'];
                $dates = '[{"op":"replace", "path":"/value", "value":'.json_encode($room['dates']).'}]';
                $product = Product::find([ 'id' => $room['id'] ])
                    ->get();

                if (!$product->isNew()) {
                    $product->scenario('update');
                    $name = 'dates';
                    $props = json_decode($product->props);
                    $prop = $props->$name;
                    $doc = json_encode($prop);
                    $patchDoc = $dates;
                    $patchedDoc = '';
                    try {
                        $patch = new Patch($doc, $patchDoc);
                        $patchedDoc = $patch->apply();
                    } catch (InvalidPatchDocumentJsonException $e) {
                        throw new Exception($e->getMessage());
                    } catch (InvalidTargetDocumentJsonException $e) {
                        throw new Exception($e->getMessage());
                    } catch (InvalidOperationException $e) {
                        throw new Exception($e->getMessage());
                    }
                    $props->$name = json_decode($patchedDoc);
                    $product->props = json_encode($props);
                } else {
                    $product->scenario('create');
                    $product->title = $title;
                    $product->id = $room['id']; 
                    $product->hash = $room['id'];
                    $product->date = time();
                    $string = '{"dates":{"type":"json","value":'.json_encode($room['dates']).'}}';
                    $product->props = $string;
                    Chpu::inject($product);
                }

                if (!$product->save()) {
                    print_r($product->getErrors());
                    throw new Exception('Продукт не сохранён');
                }
            }
        } catch (Exception $e) {
            self::sendMessageError();
            print_r($e->getMessage());
        }
    }
}
