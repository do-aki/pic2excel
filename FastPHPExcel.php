<?php
require_once dirname(__FILE__) . '/PHPExcel/Classes/PHPExcel.php';

class FastStylePHPExcel extends PHPExcel {

    private $cache = array();
	public function getCellXfByHashCode($pValue = '')
	{
        if (isset($this->cache[$pValue])) {
            return $this->cache[$pValue];
        }
        return false;
	}

	public function addCellXf(PHPExcel_Style $style)
	{
        parent::addCellXf($style);
        $this->cache[$style->getHashCode()] = $style;
	}

    public function garbageCollect() 
    {
        parent::garbageCollect();
        
        $cache = array();
        foreach ($this->getCellXfCollection() as $style) {
            $this->cache[$style->getHashCode()] = $style;
        }
    }
}
