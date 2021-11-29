<?php
/**
 *文件编码需为UTF-8，否则会存在生成的文档内容乱码
 */

/** 引入需要的类库*/
require_once '..\src\Classes\PHPExcel.php';
require_once '..\src\Classes\PHPExcel\IOFactory.php';
require_once '..\src\Classes\PHPExcel\Reader\Excel5.php';
require_once '..\src\Classes\PHPExcel\Reader\Excel2007.php';
date_default_timezone_set("Asia/Shanghai");
ob_end_clean();
$objPHPExcel = new PHPExcel();

//设置生成的Excel文件名
$date = date("Y_m_d",time());
$fileName = "{$date}.xlsx";

//测试数据，正常会从数据库中获取
$data = array(
    0 => array('id'=>2012,'name'=>'胡','age' => 25)
);

//Excel文件的说明信息
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
    ->setLastModifiedBy("Maarten Balliauw")
    ->setTitle("Office 2005 XLSX Test Document")
    ->setSubject("Office 2005 XLSX Test Document")
    ->setDescription("Test document for Office 2005 XLSX, generated using PHP classes.")
    ->setKeywords("office 2005 openxml php")
    ->setCategory("Test result file");

//设置表格内容，具体内容根据A1这种具体位置来确定
$objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('A1','编号')
    ->setCellValue('B1','姓名')
    ->setCellValue('C1','年龄');

//适合把表中数据导入Excel文件中，多数据循环设置值

foreach($data as $key=> $value) {
    $key+=2;
    $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A'.$key,$value['id'])
        ->setCellValue('B'.$key,$value['name'])
        ->setCellValue('C'.$key,$value['age']);
}
// 重命名表

// $objPHPExcel->getActiveSheet()->setTitle('Simple');

// 设置活动单指数到第一个表,所以Excel打开这是第一个表
$objPHPExcel->setActiveSheetIndex(0);

// 将输出重定向到一个客户端web浏览器(Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename='.$fileName);
header('Cache-Control: max-age=0');

//要是输出为Excel2007,使用 Excel2007对应的类，生成的文件名为.xlsx.如果是Excel2005,使用Excel5,对应生成.xls文件
//$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');

//支持浏览器下载生成的文档
$objWriter->save('php://output');




//设置excel的属性
//创建人
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw");
//最后修改人
$objPHPExcel->getProperties()->setLastModifiedBy("Maarten Balliauw");
//标题
$objPHPExcel->getProperties()->setTitle("Office 2007 XLSX Test Document");
//题目
$objPHPExcel->getProperties()->setSubject("Office 2007 XLSX Test Document");
//描述
$objPHPExcel->getProperties()->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.");
//关键字
$objPHPExcel->getProperties()->setKeywords("office 2007 openxml php");
//种类
$objPHPExcel->getProperties()->setCategory("Test result file");


//格式操作
//设置当前的sheet
$objPHPExcel->setActiveSheetIndex(0);
//设置sheet的name
$objPHPExcel->getActiveSheet()->setTitle('Simple');
//设置单元格的值
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'String');
$objPHPExcel->getActiveSheet()->setCellValue('A2', 12);
$objPHPExcel->getActiveSheet()->setCellValue('A3', true);
$objPHPExcel->getActiveSheet()->setCellValue('C5', '=SUM(C2:C4)');
$objPHPExcel->getActiveSheet()->setCellValue('B8', '=MIN(B2:C5)');
//合并单元格
$objPHPExcel->getActiveSheet()->mergeCells('A18:E22');
//分离单元格
$objPHPExcel->getActiveSheet()->unmergeCells('A28:B28');
//冻结窗口
$objPHPExcel->getActiveSheet()->freezePane('A2');
//保护cell
$objPHPExcel->getActiveSheet()->getProtection()->setSheet(true); // Needs to be set to true in order to enable any worksheet protection!
$objPHPExcel->getActiveSheet()->protectCells('A3:E13', 'PHPExcel');




//设置单元格格式

//设置格式
// Set cell number formats
echo date('H:i:s') . " Set cell number formats\n";
$objPHPExcel->getActiveSheet()->getStyle('E4')->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
$objPHPExcel->getActiveSheet()->duplicateStyle( $objPHPExcel->getActiveSheet()->getStyle('E4'), 'E5:E13' );
//设置宽width
// Set column widths
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(12);
// 设置单元格高度
// 所有单元格默认高度
$objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
// 第一行的默认高度
$objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(30);
//设置font
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setName('Candara');
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setSize(20);
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setUnderline(PHPExcel_Style_Font::UNDERLINE_SINGLE);
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objPHPExcel->getActiveSheet()->getStyle('E1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
$objPHPExcel->getActiveSheet()->getStyle('D13')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('E13')->getFont()->setBold(true);
//设置align
$objPHPExcel->getActiveSheet()->getStyle('D11')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('D12')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('D13')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
$objPHPExcel->getActiveSheet()->getStyle('A18')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY);

//垂直居中
$objPHPExcel->getActiveSheet()->getStyle('A18')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
//设置column的border
$objPHPExcel->getActiveSheet()->getStyle('A4')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
$objPHPExcel->getActiveSheet()->getStyle('B4')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
$objPHPExcel->getActiveSheet()->getStyle('C4')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
$objPHPExcel->getActiveSheet()->getStyle('D4')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
$objPHPExcel->getActiveSheet()->getStyle('E4')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
//设置border的color
$objPHPExcel->getActiveSheet()->getStyle('D13')->getBorders()->getLeft()->getColor()->setARGB('FF993300');
$objPHPExcel->getActiveSheet()->getStyle('D13')->getBorders()->getTop()->getColor()->setARGB('FF993300');
$objPHPExcel->getActiveSheet()->getStyle('D13')->getBorders()->getBottom()->getColor()->setARGB('FF993300');
$objPHPExcel->getActiveSheet()->getStyle('E13')->getBorders()->getTop()->getColor()->setARGB('FF993300');
$objPHPExcel->getActiveSheet()->getStyle('E13')->getBorders()->getBottom()->getColor()->setARGB('FF993300');
$objPHPExcel->getActiveSheet()->getStyle('E13')->getBorders()->getRight()->getColor()->setARGB('FF993300');
//设置填充颜色
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFill()->getStartColor()->setARGB('FF808080');
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFill()->getStartColor()->setARGB('FF808080');
//加图片
/*实例化插入图片类*/
$objDrawing = new PHPExcel_Worksheet_Drawing();
/*设置图片路径 切记：只能是本地图片*/
$objDrawing->setPath($img_val);
/*设置图片高度*/
$objDrawing->setWidth(200);
$img_height[] = $objDrawing->getHeight();
/*设置图片要插入的单元格*/
$objDrawing->setCoordinates($img_k[$j].$i);
/*设置图片所在单元格的格式*/
$objDrawing->setOffsetX(10);
$objDrawing->setOffsetY(10);
$objDrawing->setRotation(0);
$objDrawing->getShadow()->setVisible(true);
$objDrawing->getShadow()->setDirection(50);
$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());
//导出Excel表格例子


