<?php
namespace Trunk\ExcelLibrary\Excel;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Conditional;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class ExcelFormats {
	var $strike_through = false;
	var $data_format = null;
	var $use_border = null;

	private $dropdownOptions = false;
	public function setDropdownOptions($value) {
		$this->dropdownOptions = $value;
		return $this;
	}
	public function getDropdownOptions() {
		return $this->dropdownOptions;
	}

	private $conditionalFormatting = false;
	public function setConditionalFormatting( $value, $background ) {

		$condition = '';
		switch( substr($value, 0, 1 ) ) {
			case "!":
				$condition = Conditional::OPERATOR_NOTEQUAL;
				break;
			case "<":
				$condition = Conditional::OPERATOR_LESSTHAN;
				break;
			case ">":
				$condition = Conditional::OPERATOR_GREATERTHAN;
				break;
			case "=":
				$condition = Conditional::OPERATOR_EQUAL;
				break;
		}

		$this->conditionalFormatting = [
			'condition' => $condition,
			'match' => substr( $value, 1 ),
			'style' => [
				'font' => [
					'strike' => true,
				],
				'fill' => [
					'fillType' => Fill::FILL_SOLID,
					'fill' => Fill::FILL_SOLID,
					'color' => [
						'argb' => 'FF' . $background,
					],
					'startColor' => [
						'argb' => 'FF' . $background,
					]
				]
			],
		];
		return $this;
	}
	public function getConditionalFormatting() {
		return $this->conditionalFormatting;
	}

	private $bold = false;
	public function setBold( bool $value ) {
		$this->bold = $value;
		return $this;
	}

	private $font_colour = null;
	public function setFontColour( string $colour ) {
		$this->font_colour = $colour;
		return $this;
	}

	var $background_colour = null;
	public function setBackgroundColour( string $colour ) {
		$this->background_colour = $colour;
		return $this;
	}

	public function get_style_array() {
		$result = array();

		if ($this->strike_through) {
			$result['font']['strike'] = true;
		}

		if ($this->background_colour != null) {
			$result['fill'] = array(
				'fillType' => Fill::FILL_SOLID,
				'startColor' => array(
					'argb' => 'FF' . $this->background_colour
				)
			);
		}

		if ( $this->use_border != null ) {
			$borders_array = [];
			if ( strpos( $this->use_border, "left" ) !== false ) {
				$borders_array[ 'left' ] = [ 'style' => Border::BORDER_THIN ];
			}
			if ( strpos( $this->use_border, "right" ) !== false ) {
				$borders_array[ 'right' ] = [ 'style' => Border::BORDER_THIN ];
			}
			if ( strpos( $this->use_border, "top" ) !== false ) {
				$borders_array[ 'top' ] = [ 'style' => Border::BORDER_THIN ];
			}
			if ( strpos( $this->use_border, "bottom" ) !== false ) {
				$borders_array[ 'bottom' ] = [ 'style' => Border::BORDER_THIN ];
			}
			if ( strpos( $this->use_border, "all" ) !== false ) {
				$borders_array[ 'allborders' ] = [ 'style' => Border::BORDER_THIN ];
			}
			$result[ 'borders' ] = $borders_array;
		}

		if ( $this->font_colour ) {
			$result['font']['color'] = array( 'argb' => 'FF' . $this->font_colour );
		}
		if ( $this->bold ) {
			$result[ 'font' ][ 'bold' ] = true;
		}

		return $result;
	}
}
