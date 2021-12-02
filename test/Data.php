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
                'phone1' => 13458645501,
            ],
            [
                'id' => 2,
                'number' => 'http://www.baidu.com',
                'phone1' => 13458645501,

            ]
        ];
    }

    public function excelPath()
    {
        return __DIR__ . '/ss/'.uniqid().'.xlsx';
    }

    public function header()
    {
        return [
            'id' => "大范甘迪",
            'number' => "编号",
            'phone1' => "电话2",
        ];
    }
}