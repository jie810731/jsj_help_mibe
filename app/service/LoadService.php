<?php

namespace App\Service;

use Illuminate\Support\Facades\Storage;

class LoadService
{
    public function getSpreadSheetData($spreadsheet)
    {
        $data = [];
        $sheet_names = $spreadsheet->getSheetNames();

        foreach ($sheet_names as $key => $sheet_name) {
            $rows = [];
            foreach ($spreadsheet->getSheet($key)->getRowIterator() as $row) {
                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(false);
                $cells = [];
                foreach ($cellIterator as $cell) {
                    $cell_value = $cell->getValue();
                    $cell_value = trim($cell_value);
                    if (\PhpOffice\PhpSpreadsheet\Shared\Date::isDateTime($cell)) {
                        $date_object = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($cell_value);
                        $cell_value = $date_object->format('Y/m/d');
                    }
                    $cells[] = $cell_value;
                }
                $rows[] = $cells;
            }

            $data[] = $rows;
        }

        return $data;
    }
}
