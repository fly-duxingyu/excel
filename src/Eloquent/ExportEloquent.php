<?php

namespace Duxingyu\Excel\Eloquent;

use Duxingyu\Excel\Contracts\ExcelHeaderInterface;
use Duxingyu\Excel\Contracts\QueryDataInterface;
use Duxingyu\Excel\Contracts\StorageDataInterface;
use PHPExcel;
use PHPExcel_IOFactory;

/**
 *导出
 */
abstract class ExportEloquent implements QueryDataInterface, StorageDataInterface, ExcelHeaderInterface
{
    public $letter = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

    /**
     * 处理导出
     * @return array
     */
    public function executeExportData()
    {
        try {
            $data = $this->getData();
            if (empty($data)) {
                throw new \Exception('导出数据为空,无需执行任务');
            }
            $header = $this->header();
            $path = $this->excelPath();
            array_unshift($data, $header);
            $data = array_filter($data);
            sort($data);
            $objPHPExcel = new PHPExcel();
            $keyIndex = count(reset($data));
            foreach ($data as $key => $value) {
                $key += 1;
                $value = array_values($value);
                for ($i = 0; $i < $keyIndex; $i++) {
                    $objPHPExcel->setActiveSheetIndex(0)->setCellValue($this->letter[$i] . $key, $value[$i]);
                }
            }
            $objPHPExcel->getActiveSheet()->setTitle('导出');
            $objPHPExcel->setActiveSheetIndex(0);
            ob_end_clean();//清除缓冲区,避免乱码
            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
            $objWriter->save($path);
            return [
                'status' => 'ok',
                'message' => '导出成功',
                'detail_message' => '',
                'path' => $path
            ];
        } catch (\Exception  $exception) {
            return [
                'status' => 'no',
                'message' => '导出失败',
                'detail_message' => $exception->getMessage(),
                'path' => $this->excelPath()
            ];

        }
    }
}