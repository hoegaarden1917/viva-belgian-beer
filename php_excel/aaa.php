<?php

set_include_path(get_include_path() . PATH_SEPARATOR . './PHPExcel/Classes/');

//配列作成
$cd_list = array(
    array('(3rd Album)', '2009', '07', '08', 'AL'),
    array('ワンルーム・ディスコ', '2009', '03', '25', 'SG'),
    array('Dream Fighter', '2008', '11', '19', 'SG'),
    array('love the world', '2008', '07', '09', 'SG'),
    array('GAME', '2008', '04', '16', 'AL'),
);

//セルの書式 (文字列、上下左右に罫線)
$cell_style = array(
    'borders' => array(
        'top'     => array('style' => PHPExcel_Style_Border::BORDER_THIN),
        'bottom'  => array('style' => PHPExcel_Style_Border::BORDER_THIN),
        'left'    => array('style' => PHPExcel_Style_Border::BORDER_THIN),
        'right'   => array('style' => PHPExcel_Style_Border::BORDER_THIN)
    )
);



include 'PHPExcel.php';
include 'PHPExcel/IOFactory.php';

//テンプレートを読み込んでインスタンス化
$objReader = PHPExcel_IOFactory::createReader("Excel5");
$objPHPExcel = $objReader->load("./template.xls");
            
//データのセット
$row = 3;		//ヘッダの行数を除く
foreach ($cd_list as $cd) {
    $col = 0;
    foreach ($cd as $value) {
        $objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow($col, $row, $value, PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->getStyleByColumnAndRow($col++, $row)->applyFromArray($cell_style);
    }
    $row++;
}

//Excel2003以前の形式でファイル出力
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('./cd_list.xls');
?>