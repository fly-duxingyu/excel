<?php

use Duxingyu\Excel\Eloquent\ExportEloquent;


class Down extends ExportEloquent
{
    protected $data;
    protected $header;
    protected $path;

    public function __construct($data, $header, $path)
    {
        $this->data = $data;
        $this->header = $header;
        $this->path = $path;
    }

    public function getData()
    {
        return $this->data;
    }

    public function excelPath()
    {
        return $this->path;
    }

    public function header()
    {
        return $this->header;
    }
}