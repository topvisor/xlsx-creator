<?php

namespace Decaseal\XlsxCreator\Xml\Style\Index;

use Decaseal\XlsxCreator\Xml\BaseXml;
use Iterator;

class Index implements Iterator{
	private $indexes;
	private $position;

	public function __construct(){
		$indexes = [];
		$position = 0;
	}

	function addIndex(BaseXml $baseXml, $model = null) : int{
		$xml = $baseXml->toXml($model);

		$index = $this->indexes[$xml] ?? false;

		if($index === false){
			$index = count($this->indexes);
			$this->indexes[$xml] = $index;
		}

		return $index;
	}

	public function current(){
		return array_keys($this->indexes)[$this->position];
	}

	public function next(){
		$this->position++;
	}

	public function key(){
		return $this->indexes[$this->current()];
	}

	public function valid(){
		return $this->indexes && count(array_keys($this->indexes)) > $this->position;
	}

	public function rewind(){
		$this->position = 0;
	}
}