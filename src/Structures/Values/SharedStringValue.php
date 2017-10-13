<?php

namespace Topvisor\XlsxCreator\Structures\Values;

use Topvisor\XlsxCreator\Helpers\Validator;

class SharedStringValue extends Value{
	function __construct(int $id){
		Validator::validatePositive($id, '$id');

		parent::__construct($id, Value::TYPE_SHARED_STRING);
	}

	static function parse($value): Value{
		return new self($value);
	}
}