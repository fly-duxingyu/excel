<?php

use Duxingyu\Excel\Eloquent\ImportEloquent;


class Data1 extends ImportEloquent
{
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