<?php

namespace Duxingyu\Excel\Eloquent;

use Duxingyu\Excel\Contracts\QueryDataInterface;
use Duxingyu\Excel\Contracts\StorageDataInterface;
use PHPExcel;
use PHPExcel_IOFactory;

/**
 *导入
 */
abstract class ImportEloquent implements QueryDataInterface, StorageDataInterface
{
    public function executeData()
    {
        try {
            $data = $this->getData();
            $header = $this->getHeader();
            $excelName = $this->getExcelName();
            $path = $this->getPath();
            array_unshift($data, $header);
            $date = "_" . date("Ymd") . uniqid();
            $fileName = $excelName . $date . ".xlsx";
            $objPHPExcel = new PHPExcel();
            $keyIndex = count(reset($data));
            foreach ($data as $key => $value) {
                $key += 1;
                $value = array_values($value);
                for ($i = 0; $i < $keyIndex; $i++) {
                    $objPHPExcel->setActiveSheetIndex(0)->setCellValue($this->letter[$i] . $key, $value[$i]);
                    $objPHPExcel->setActiveSheetIndex(0)->setCellValue($this->letter[$i] . $key, $value[$i]);
                    $objPHPExcel->setActiveSheetIndex(0)->setCellValue($this->letter[$i] . $key, $value[$i]);
                    $objPHPExcel->setActiveSheetIndex(0)->setCellValue($this->letter[$i] . $key, $value[$i]);
                }
            }
            $objPHPExcel->getActiveSheet()->setTitle('导出');
            $objPHPExcel->setActiveSheetIndex(0);
            ob_end_clean();//清除缓冲区,避免乱码
            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
            $pathName = $path . "\\" . $fileName;
            $objWriter->save($pathName);
            return [
                'status' => 'ok',
                'message' => '导出成功',
                'data' => [
                    'full_path' => $pathName,
                    'file_name' => $fileName,
                    'path' => $path,
                ]];
        } catch (\Exception  $exception) {
            return [
                'status' => 'no',
                'message' => $exception->getMessage(),
                'data' => [
                    'full_path' => '',
                    'file_name' => '',
                    'path' => '',
                ]];
        }
    }
}