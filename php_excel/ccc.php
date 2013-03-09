<?php
 // 必要なクラスをインクルードする
set_include_path(get_include_path() . PATH_SEPARATOR . "./PHPExcel/Classes/");
include "PHPExcel.php";
include "PHPExcel/IOFactory.php";

// PHPExcelオブジェクトを生成する
// $reader = PHPExcel_IOFactory::createReader("Excel2007");
$reader = PHPExcel_IOFactory::createReader("Excel5");
$book = $reader->load("./template.xls");

// シートの設定を行う
$book->setActiveSheetIndex(0);
$sheet = $book->getActiveSheet();

// セルに値をセットする
$sheet->setCellValue("C2", "山田 太郎");
$sheet->setCellValue("F2", "2009/3/22");

// Excel2007形式で保存する
//$writer = PHPExcel_IOFactory::createWriter($book, "Excel2007");
$writer = PHPExcel_IOFactory::createWriter($book, "Excel5");
$writer->save("./output.xls"); 
?>
