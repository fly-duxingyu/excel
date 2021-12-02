<?php

namespace Duxingyu\Excel\Contracts;

/**
 * 获取数据接口
 */
interface QueryDataInterface
{
    /**
     * 获取导出数据
     * @return mixed
     */
    public function getData();
}