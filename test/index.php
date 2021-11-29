<?php

use Duxingyu\Excel\Eloquent\ExcelEloquent;

require __DIR__.'/../vendor/autoload.php';
require __DIR__.'/Data.php';
require __DIR__.'/Data1.php';
$data=(new ExcelEloquent(new Data))->execute();
$data=(new ExcelEloquent(new Data1))->execute();
echo '<pre>';
print_r($data);
