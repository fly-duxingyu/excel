<?php

namespace Duxingyu\Excel\Contracts;

interface  StorageDataInterface
{
    /**
     * 设置导入导出的文件名称
     * @return mixed
     */
    public function setExcelName();

    /**
     * 设置文件路径
     * @return mixed
     */
    public function setPath();
}