<?php

namespace Duxingyu\Excel\Eloquent;

require_once __DIR__ . '\..\Classes\PHPExcel.php';
require_once __DIR__ . '\..\Classes\PHPExcel\IOFactory.php';
require_once __DIR__ . '\..\Classes\PHPExcel\Reader\Excel5.php';
require_once __DIR__ . '\..\Classes\PHPExcel\Reader\Excel2007.php';

use Duxingyu\Excel\Contracts\QueryDataInterface;

/**
 * 工厂类
 */
class ExcelEloquent
{
    /**
     * @var QueryDataInterface
     */
    protected $object;

    public function __construct(QueryDataInterface $disposeData)
    {
        $this->object = $disposeData;
    }

    public function execute()
    {
        return $this->object->executeData();
    }
}