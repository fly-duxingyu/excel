<?php

namespace Duxingyu\Excel\Contracts;

/**
 * 处理数据
 */
interface DisposeDataInterface
{
    /**
     * 错误处理
     * @return mixed
     */
    public function errorDealData();
}