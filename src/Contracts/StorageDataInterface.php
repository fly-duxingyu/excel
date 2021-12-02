<?php

namespace Duxingyu\Excel\Contracts;

interface  StorageDataInterface
{
    /**
     * 设置导入导出的文件路径地址
     * @return mixed
     */
    public function excelPath();
}