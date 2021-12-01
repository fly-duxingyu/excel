<?php

namespace Duxingyu\Excel\Eloquent;

use Duxingyu\Excel\Contracts\QueryDataInterface;
use Duxingyu\Excel\Contracts\StorageDataInterface;
use PHPExcel;
use PHPExcel_Cell;
use PHPExcel_IOFactory;

/**
 *导入
 */
abstract class ImportEloquent implements QueryDataInterface, StorageDataInterface
{
    /**
     * 设置导出数据头
     * @return array
     */
    abstract protected function setHeader(): array;

    /**
     * 获取excel头部
     * @return array
     */
    protected function getHeader()
    {
        return $this->setHeader();
    }

    public function getData()
    {
        $header = $this->getHeader();
        $header = array_flip($header);
        $reader = PHPExcel_IOFactory::createReader('Excel2007'); //设置以Excel5格式(Excel97-2003工作簿)
        $PHPExcel = $reader->load("D:/phpstudy_pro/WWW/excel/test/excel.xlsx"); // 载入excel文件
        $sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
        $highestRow = $sheet->getHighestRow(); // 取得总行数
        $highestColumm = $sheet->getHighestColumn(); // 取得总列数
        $highestColumm = PHPExcel_Cell::columnIndexFromString($highestColumm); //字母列转换为数字列 如:AA变为27

        $initData = [];
        $key = [];
        /** 循环读取每个单元格的数据 */
        for ($row = 1; $row <= $highestRow; $row++) {//行数是以第1行开始
            $data = [];
            for ($column = 0; $column < $highestColumm; $column++) {//列数是以第0列开始
                $columnName = PHPExcel_Cell::stringFromColumnIndex($column);
                if ($row == 1) {
                    array_push($key, $sheet->getCellByColumnAndRow($column, $row)->getValue());
                    continue;
                }
                $data = array_merge($data, [$header[$key[$column]] => $sheet->getCellByColumnAndRow($column, $row)->getValue()]);
//                echo $sheet->getCellByColumnAndRow($column, $row)->getValue()."<br />";
            }
            if (!empty($data)) {
                $initData[] = $data;
            }

        }
        return $initData;
    }

    public function executeData()
    {
        echo md5('scjpf'.'cKPpgoQcfYcN1sMrTPich6wwArxJLi7F1638348748');die();
        echo '<pre>';
        try {
            $data = $this->getData();
            print_r($data);die();
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