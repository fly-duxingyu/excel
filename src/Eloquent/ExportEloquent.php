<?php

namespace Duxingyu\Excel\Eloquent;

use Duxingyu\Excel\Contracts\ExcelHeaderInterface;
use Duxingyu\Excel\Contracts\QueryDataInterface;
use Duxingyu\Excel\Contracts\StorageDataInterface;
use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Style_Alignment;

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
            if (!empty($header)) {
                array_unshift($data, $header);
            }
            $data = array_filter($data);
            $objPHPExcel = new PHPExcel();
            $keyIndex = count(reset($data));
            foreach ($data as $key => $value) {
                $key += 1;
                $value = array_values($value);
                for ($i = 0; $i < $keyIndex; $i++) {
                    $url = parse_url($value[$i]);
                    if (!empty($url['scheme'])) {
                        $obj = $objPHPExcel->setActiveSheetIndex()->setCellValue($this->letter[$i] . $key, $value[$i]);
                        //设置单元格超链接
                        $obj->getCell($this->letter[$i] . $key)->getHyperlink()->setUrl($value[$i]);
                        //设置单元格样式
                        $obj->getStyle($this->letter[$i] . $key)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                        continue;
                    }
                    //设置值和样式
                    $objPHPExcel->setActiveSheetIndex(0)->setCellValue($this->letter[$i] . $key, $value[$i])->getStyle($this->letter[$i] . $key)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                }
            }
            $objPHPExcel->getActiveSheet()->setTitle('导出');
            $objPHPExcel->setActiveSheetIndex(0);
            ob_end_clean();//清除缓冲区,避免乱码
            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
            if (file_exists($path)) {
                $path = mb_substr($path, 0, strripos($path, '.')) . uniqid() . '.xlsx';
            }
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