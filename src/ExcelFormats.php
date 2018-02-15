<?php
namespace Trunk\ExcelLibrary\Excel;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class ExcelFormats {
	var $font_colour = null;
	var $background_colour = null;
	var $strike_through = false;
	var $data_format = null;
	var $use_border = null;

	public function get_style_array() {
		$result = array();

		if ($this->strike_through) {
			$result['font']['strike'] = true;
		}

		if ($this->background_colour != null) {
			$result['fill'] = array(
				'type' => Fill::FILL_SOLID,
				'startcolor' => array(
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

		return $result;
	}
}
