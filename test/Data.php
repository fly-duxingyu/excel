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

    protected function setHeader(): array
    {
        return [
            'id' => "ID",
            'number' => "编号",
            'name' => "姓名",
            'phone' => "电话",
        ];
    }

    public function setExcelName()
    {
        return '测试文件';
    }

    public function setPath()
    {
        return __DIR__ . '\ss';
    }
}