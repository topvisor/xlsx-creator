<?php

namespace Topvisor\XlsxCreator\Structures;

use Topvisor\XlsxCreator\Exceptions\InvalidValueException;
use Topvisor\XlsxCreator\Helpers\Serializable;
use Topvisor\XlsxCreator\Helpers\Validator;

/**
 * Class Color. Используется для задания цвета.
 *
 * @package  Topvisor\XlsxCreator
 */
class Color implements Serializable {
	private $model;

	/**
	 * Color constructor.
	 * @param array $model - модель цвета
	 */
	private function __construct(array $model) {
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
	public static function fromHex(string $rgb = 'FFFFFF', string $a = 'FF'): self {
		switch (mb_strlen($rgb)) {
			case 3: $rgb = preg_replace('/./', '$0$0', $rgb);

			break;
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
	public static function fromInt(int $r = 0, int $g = 0, int $b = 0, int $a = 255): self {
		Validator::validateInRange($r, 0, 255, '$r');
		Validator::validateInRange($g, 0, 255, '$g');
		Validator::validateInRange($b, 0, 255, '$b');
		Validator::validateInRange($a, 0, 255, '$a');

		return new self(['argb' => dechex($a) . dechex($r) . dechex($g) . dechex($b)]);
	}

	/**
	 * Задать цвет из предустановленных
	 *
	 * @param int $theme - номер цвета
	 * @return Color - цвет
	 * @throws InvalidValueException
	 */
	public static function fromTheme(int $theme): self {
		Validator::validatePositive($theme, '$theme');

		return new self(['theme' => $theme]);
	}

	/**
	 * @return array - модель цвета
	 */
	public function getModel(): array {
		return $this->model;
	}

	public function serialize() {
		return $this->model['argb'] ? $this->model['argb'] : (string) $this->model['theme'];
	}

	public function unserialize($serialized) {
		if (preg_match('/^([A-F\d]{2})([A-F\d]{6})$/', $serialized, $matches))
			$this->model = self::fromHex($matches[2], $matches[1])->getModel();
			else $this->model = self::fromTheme((int) $serialized)->getModel();
	}
}
