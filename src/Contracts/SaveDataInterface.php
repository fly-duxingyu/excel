<?php

namespace Duxingyu\Excel\Contracts;

interface SaveDataInterface
{
    /**
     * 保存数据
     * @param $data //需要保存的数据
     * @return mixed
     */
    public function saveData($data);

    /**
     * 验证导入数据
     * @param $data //初始数据
     * @param $correctData //正确通过验证数据
     * @param $errorData //错误未通过验证数据
     * @return mixed
     */
    public function checkData($data, &$correctData, &$errorData);
}