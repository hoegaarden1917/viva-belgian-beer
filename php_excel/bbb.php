<?php

set_include_path(get_include_path() . PATH_SEPARATOR . './PHPExcel/Classes/');

include_once 'PHPExcel.php';
include_once 'PHPExcel/IOFactory.php';

// Excel 97-2003 形式で作成し、ブラウザで開く
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="excel97-2003.xls"');
header('Cache-Control: max-age=0');

$excel = new PHPExcel();
// 1番目のシートに書き込みを行う
$excel->setActiveSheetIndex(0);
$excel->getActiveSheet()->setCellValue('A1', 'PHPExcel 動作確認');

// Excel2007形式の場合は、第2パラメータに 'Excel2007' を指定する
$writer = PHPExcel_IOFactory::createWriter($excel, 'Excel5');
// ブラウザに出力
//$writer->save('php://output');

// ファイルに出力する場合は、パスを指定すし、前述の header 関数は全て削除する
$writer->save('./excel.xls');

?>