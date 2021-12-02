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

    /**
     * 执行导出
     * @return mixed
     */
    public function downExcel()
    {
        return $this->object->executeExportData();
    }

    /**
     * 执行导入
     * @return mixed
     */
    public function importExcel()
    {
        return $this->object->executeImportData();
    }
}