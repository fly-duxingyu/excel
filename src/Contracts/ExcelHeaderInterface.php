<?php

namespace Duxingyu\Excel\Contracts;

interface ExcelHeaderInterface
{
    /**
     * 导入导出的header头部
     * @return mixed
     */
    public function header();
}