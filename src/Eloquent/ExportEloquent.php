<?php

namespace Duxingyu\Excel\Eloquent;

use Duxingyu\Excel\Contracts\QueryDataInterface;
use Duxingyu\Excel\Contracts\StorageDataInterface;
use PHPExcel;
use PHPExcel_IOFactory;

/**
 *导出
 */
abstract class ExportEloquent implements QueryDataInterface, StorageDataInterface
{
    //获取数据
    //处理数据
    //错误异常处理  下载地址
    //消息写入数据库操作
    public $letter = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

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

    /**
     * 获取excel名称
     * @return array
     */
    protected function getExcelName()
    {
        return $this->setExcelName();
    }

    /**
     * 获取文件路径
     * @return mixed
     */
    protected function getPath()
    {
        return $this->setPath();
    }

    /**
     * 处理导出
     * @return array
     */
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