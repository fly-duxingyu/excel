<?php

use Duxingyu\Excel\Eloquent\ImportEloquent;


class Data1 extends ImportEloquent
{
    public function header()
    {
        return [
            'id' => "ID",
            'number' => "编号",
            'name' => "姓名",
            'phone' => "电话",
        ];
    }

    public function excelPath()
    {
        return 'D:/phpstudy_pro/WWW/excel/test/excel.xlsx';
    }

    public function saveData($data)
    {
    }

    public function checkData($data, &$correctData, &$errorData)
    {
        foreach ($data as $item) {
            if (empty($item)) {
                $errorData[] = $item;
                continue;
            }
            if (!is_numeric($item['id'])) {
                $item['error_message']='id必须是数字';
                $errorData[] = $item;
                continue;
            }
            $correctData[] = $item;
        }
    }
}