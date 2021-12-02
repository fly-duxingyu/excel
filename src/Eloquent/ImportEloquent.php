<?php

namespace Duxingyu\Excel\Eloquent;

use Duxingyu\Excel\Contracts\ExcelHeaderInterface;
use Duxingyu\Excel\Contracts\QueryDataInterface;
use Duxingyu\Excel\Contracts\SaveDataInterface;
use Duxingyu\Excel\Contracts\StorageDataInterface;
use PHPExcel;
use PHPExcel_Cell;
use PHPExcel_Exception;
use PHPExcel_IOFactory;
use PHPExcel_Reader_Exception;
use PHPExcel_Style_Alignment;
use PHPExcel_Writer_Exception;

/**
 *导入
 */
abstract class ImportEloquent implements QueryDataInterface, StorageDataInterface, ExcelHeaderInterface, SaveDataInterface
{
    public $letter = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

    /**
     * 读取excel表格
     * @return array
     * @throws PHPExcel_Exception
     * @throws PHPExcel_Reader_Exception
     */
    final public function getData()
    {
        $path = $this->excelPath();
        if (!file_exists($path)) {
            throw new \Exception('导入文件地址不存在');
        }
        $header = $this->header();
        $header = array_flip($header);
        $reader = PHPExcel_IOFactory::createReader('Excel2007'); //设置以Excel5格式(Excel97-2003工作簿)
        $PHPExcel = $reader->load($path); // 载入excel文件
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

    /**
     * 写入错误数据
     * @param $errorData
     * @return string
     * @throws PHPExcel_Exception
     * @throws PHPExcel_Reader_Exception
     * @throws PHPExcel_Writer_Exception
     */
    final protected function writeExcel($errorData)
    {
        $header = $this->header();
        $header['error_message'] = '错误描述';
        $path = $this->excelPath();
        $path = mb_substr($path, 0, strripos($path, '.')) . '错误数据.xlsx';
        array_unshift($errorData, $header);
        $objPHPExcel = new PHPExcel();
        $keyIndex = count(reset($errorData));
        foreach ($errorData as $key => $value) {
            $key += 1;
            $value = array_values($value);
            for ($i = 0; $i < $keyIndex; $i++) {
                $objPHPExcel->setActiveSheetIndex(0)->setCellValue($this->letter[$i] . $key, $value[$i])->getStyle($this->letter[$i] . $key)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            }
        }
        $objPHPExcel->getActiveSheet()->setTitle('错误数据');
        $objPHPExcel->setActiveSheetIndex(0);
        ob_end_clean();//清除缓冲区,避免乱码
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        if (file_exists($path)) {
            $path = mb_substr($this->excelPath(), 0, strripos($this->excelPath(), '.')) . '错误数据' . uniqid() . '.xlsx';
        }
        $objWriter->save($path);
        return $path;
    }

    /**
     * 导入执行逻辑
     * @return array
     */
    final public function executeImportData()
    {
        try {
            //读取excel数据
            $data = $this->getData();
            $this->checkData($data, $correctData, $errorData);
            $path = $this->writeExcel($errorData);
            $this->saveData($correctData);
            return [
                'status' => 'ok',
                'message' => '导入成功',
                'detail_message' => '',
                'path' => $path
            ];
        } catch (\Exception  $exception) {
            return [
                'status' => 'no',
                'message' => '导入失败',
                'detail_message' => $exception->getMessage(),
                'path' => $this->excelPath()
            ];
        }
    }
}