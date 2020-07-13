<?php
require "vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer;

//Константы
const maxCol = 9;
const maxPage = 2;

//Переменные состояния
$pageIndex = 0;
$emptyRows = 0;
$emptyCols = 0;
$isRowEmpty = false;
$isSectorEmpty = false;

//Переменные с данными
$cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"];
$usageFile = "PRAJS_DAIKIN_dlya_sayta.xlsx";

//Объекты
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
$workSpreadsheet = new Spreadsheet();

//Переменные со значениями
$row = 11;
$nextRow = 11;
$stape = -10;
$id = 0;
$rowForSpreadsheet = 11;

//Шаблон аттрибутов
$attributes = [["Основные характеристики","Внутренний блок"],
               ["Основные характеристики","Наружный блок"],
               ["Основные характеристики","Пульт дистанционного управления"],
               ["Основные характеристики","Панель"],
               ["Основные характеристики","Тип кондиционера"],
               ["Основные характеристики","фреон"],
               ["Основные характеристики","Произв., кВт холод"],
               ["Основные характеристики","Произв., кВт тепло"],
               ["Основные характеристики","Обслуживаемая площадь","-"],
               ["Основные характеристики","Максимальная длина коммуникаций", "-"],
               ["Основные характеристики","Класс энергопотребления", "-"],
               ["Основные характеристики","Основные режимы", "-"],
               ["Габариты","Внутреннего блока сплит-системы или мобильного кондиционера", "-"],
               ["Габариты","Вес внутреннего блока", "-"]];


$spreadsheet = $reader -> load($usageFile); 
$spreadsheet -> setActiveSheetIndex($pageIndex);

while($emptyRows <= 10){

    for ($col = 2; $col <= maxCol; $col++){
        $cord = (string)$cols[$col].$row;
        $attribute = (string)$spreadsheet -> getActiveSheet() -> getCell($cord);

        if ($attribute == null || $attribute == "0"){
            $attribute = "-";
            $emptyCols ++;
            $isSectorEmpty = true;
        }
        else{
            $emptyCols = 0;
            $isSectorEmpty = false;
        }

        if ($emptyCols >= 3){
            $isRowEmpty = true;
            break;
        }
        else{
            $isRowEmpty = false;
        }
        $attributes[$col-2][2] = $attribute;
    }

    if ($isRowEmpty){
        $emptyRows++;
        if ($emptyRows >= 9 && $pageIndex < maxPage){
                $pageIndex ++;
                $spreadsheet -> setActiveSheetIndex($pageIndex);
                $row = 11;
                continue;
        }
    }
    else{
        $id ++;
        $emptyRows = 0;
        foreach ($attributes as $attr){
            $workSpreadsheet -> getActiveSheet() 
                             -> setCellValue("A".($rowForSpreadsheet + $stape), $attr[0])
                             -> setCellValue("B".($rowForSpreadsheet + $stape), $attr[1])
                             -> setCellValue("C".($rowForSpreadsheet + $stape), $attr[2])
                             -> setCellValue("D".($rowForSpreadsheet + $stape), $id);
            $stape++;
        }
        $stape += 2;
    }
    $rowForSpreadsheet ++;
    $row ++;

}

$writer = new Writer\Xlsx ($workSpreadsheet);
$writer -> save("productAttribute.xlsx");

?>