<?php

use Duxingyu\Excel\Eloquent\ExcelEloquent;

require __DIR__ . '/../vendor/autoload.php';
require __DIR__ . '/Data.php';
require __DIR__ . '/Data1.php';

require __DIR__ . '/Down.php';


//导出
$data = (new ExcelEloquent(new Data))->downExcel();

//导入
$data = (new ExcelEloquent(new Data1))->importExcel();


//可实现脚本定时执行,有数据在调用导出类
$data = [
    [
        'id' => 1,
        'number' => 'Act123456',
        'phone' => 13458645501,
        'sex' => '男',
    ],
    [
        'id' => 2,
        'number' => 'http://www.baidu.com',
        'phone' => 13458645501,
        'sex' => '男',
    ]
];
$header = [
    'id' => "ID",
    'number' => "编号",
    'phone' => "电话",
    'sex' => "性别",
];
$path = __DIR__ . '/ss/' . uniqid() . '.xlsx';
$data = (new ExcelEloquent(new Down($data, $header, $path)))->downExcel();
