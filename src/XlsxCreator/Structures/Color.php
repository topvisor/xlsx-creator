<?php

namespace XlsxCreator\Structures;

use XlsxCreator\Exceptions\InvalidValueException;

/**
 * Class Color. Используется для задания цвета.
 *
 * @package XlsxCreator
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
		self::validateHex($rgb, '$rgb');
		self::validateHex($a, '$a');

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
		self::validateInt($r, '$r');
		self::validateInt($g, '$g');
		self::validateInt($b, '$b');
		self::validateInt($a, '$a');

		return new self(['argb' => dechex($a) . dechex($r) . dechex($g) . dechex($b)]);
	}

	/**
	 * @return array - модель цвета
	 */
	function getModel(): array{
		return $this->model;
	}

	/**
	 * @param string $hex - проверяемое значение
	 * @param string $varName - название параметра (для сообщения об ошибке)
	 * @throws InvalidValueException
	 */
	private static function validateHex(string $hex, string $varName){
		if (preg_match('/[^\dA-F]/i', $hex, $matches))
			throw new InvalidValueException("Invalid character '$matches[0]' in $varName: $varName must be hex");
	}

	/**
	 * @param int $int - проверяемое значение
	 * @param string $varName - название параметра (для сообщения об ошибке)
	 * @throws InvalidValueException
	 */
	private static function validateInt(int $int, string $varName){
		if ($int < 0 || $int > 255)	throw new InvalidValueException("$varName must be in [0;255]");
	}
}