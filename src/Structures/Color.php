<?php

namespace Topvisor\XlsxCreator\Structures;

use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Validator;

/**
 * Class Color. Используется для задания цвета.
 *
 * @package  Topvisor\XlsxCreator
 */
class Color{
	private $model;

	/**
	 * Color constructor.
	 * @param array $model - модель цвета
	 */
	private function __construct(array $model){
		$this->model = $model;
	}

	/**
	 * Задать цвет hex строкой
	 *
	 * @param string $rgb - rgb
	 * @param string $a - прозрачность
	 * @return Color - цвет
	 * @throws InvalidValueException
	 */
	static function fromHex(string $rgb = 'FFFFFF', string $a = 'FF') : self{
		switch (mb_strlen($rgb)) {
			case 3: $rgb = preg_replace('/./', '$0$0', $rgb); break;
			case 6: break;
			default: throw new InvalidValueException('The length $rgb must be 3 or 6');
		}

		if (mb_strlen($a) !== 2) throw new InvalidValueException('The length $a must be 2');
		Validator::validateHex($rgb, '$rgb');
		Validator::validateHex($a, '$a');

		return new self(['argb' => mb_strtoupper($a) . mb_strtoupper($rgb)]);
	}

	/**
	 * Задать цвет с помощью int
	 *
	 * @param int $r - красный
	 * @param int $g - зеленый
	 * @param int $b - синий
	 * @param int $a - прозрачность
	 * @return Color - цвет
	 * @throws InvalidValueException
	 */
	static function fromInt(int $r = 0, int $g = 0, int $b = 0, int $a = 255) : self{
		Validator::validateInRange($r, 0, 255, '$r');
		Validator::validateInRange($g, 0, 255, '$g');
		Validator::validateInRange($b, 0, 255, '$b');
		Validator::validateInRange($a, 0, 255, '$a');

		return new self(['argb' => dechex($a) . dechex($r) . dechex($g) . dechex($b)]);
	}

	/**
	 * @return array - модель цвета
	 */
	function getModel(): array{
		return $this->model;
	}
}