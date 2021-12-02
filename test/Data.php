<?php

use Duxingyu\Excel\Eloquent\ExportEloquent;


class Data extends ExportEloquent
{
    public function getData()
    {
        return [
            [
                'id' => 1,
                'number' => 'Act123456',
                'name' => '张三',
                'phone' => 13458645501,
            ],
            [
                'id' => 2,
                'number' => 'Act11123456',
                'name' => '张1三',
                'phone' => 13458645502,
            ]
        ];
    }

    public function excelPath()
    {
        return __DIR__ . '/ss/1.xlsx';
    }

    public function header()
    {
        return [
            'id' => "ID",
            'number' => "编号",
            'name' => "姓名",
            'phone' => "电话",
        ];
    }
}