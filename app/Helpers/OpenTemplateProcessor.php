<?php

namespace App\Helpers;

use PhpOffice\PhpWord\TemplateProcessor;

class OpenTemplateProcessor extends TemplateProcessor
{
    protected $_instance;

    public function __construct($instance) 
    {
		return parent::__construct($instance);
    }

    public function __get($key) 
    {
        return $this->$key;
    }

    public function __set($key, $val) 
    {
        return $this->$key = $val;
    }
}