<?php
namespace Trunk\ExcelLibrary\Excel;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class ExcelSheet {
//https://github.com/PHPOffice/PHPExcel/blob/develop/Documentation/markdown/Overview/08-Recipes.md
//https://github.com/PHPOffice/PHPExcel/blob/develop/Documentation/markdown/Overview/11-Appendices.md

	var $tabTitle;
	var $sheetTitle;
	var $sheetDescription;

	/**
	 * @var array Holds the actual body data for the excel sheet
	 */
	var $data;

	/**
	 * @var array
	 */
	var $header_rows;

	var $num_header_rows = 1;

	var $formatMap;

	var $row_formats = array();
	var $cell_formats = array();

	var $end_column_number = 0;
	var $end_column_letter = null;

	/**
	 * Gets the last column letter for the given data
	 */
	public function get_end_column_letter() {

		if ($this->end_column_letter != null)
			return $this->end_column_letter;

		// If the number of columns is greater than 26
		/*if ( $this->end_column_number > 26) {
			$num_round_alphabet = floor($this->end_column_number / 26);
			$letter_number = $this->end_column_number - ( $num_round_alphabet * 26 );
			$this->end_column_letter = chr( $letter_number + 64 );
			$this->end_column_letter = chr( $num_round_alphabet + 64 ) . $this->end_column_letter;
		}
		else {
			$this->end_column_letter = chr( $this->end_column_number + 64 );
		}*/

		$this->end_column_letter = Coordinate::stringFromColumnIndex($this->end_column_number);

		return $this->end_column_letter;
	}

	public static function get_letter_from_number( $n ) {
		for($r = ""; $n >= 0; $n = intval($n / 26) - 1)
			$r = chr($n%26 + 0x41) . $r;
		return $r;
	}

	/**
	 * Adds formatting for the given row, read from the given element (normally a tr)
	 * @param type $row_number
	 * @param type $element
	 * @return type
	 */
	public function add_row_format($row_number, $element) {
		$format = $this->_get_formatting($element);

		if ($format == null)
			return;

		$this->row_formats[$row_number] = $format;
	}

	/**
	 * Adds formatting for the given cell, read from the given element (normally a td)
	 * @param type $cell_key
	 * @param type $element
	 * @return type
	 */
	public function add_cell_format($cell_key, $element) {
		$format = $this->_get_formatting($element);

		if ($format == null)
			return;

		$this->cell_formats[$cell_key] = $format;
	}

	/**
	 * Set the format on the given cell for this row
	 * @param type $row_number
	 * @param type $cell
	 */
	public function set_row_format($active_sheet, $excel_row_number, $row_number) {
		if (!isset($this->row_formats[$row_number]))
			return;

		$row_range = 'A' . $excel_row_number . ':' . $this->get_end_column_letter() . $excel_row_number;

		$active_sheet->getStyle($row_range)->applyFromArray($this->row_formats[$row_number]->get_style_array());
	}

	/**
	 * Set the format on the given cell for this cell
	 * @param type $row_number
	 * @param type $cell
	 */
	public function set_cell_format($active_sheet, $cell_reference, $cell_key) {
		if (!isset($this->cell_formats[$cell_key]))
			return;

		$active_sheet->getStyle($cell_reference)->applyFromArray($this->cell_formats[$cell_key]->get_style_array());

		$data_format = $this->cell_formats[$cell_key]->data_format;
		switch ($data_format) {
			case "currency":
				$active_sheet->getStyle($cell_reference)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);
				break;
			case "date":
				$active_sheet->getStyle($cell_reference)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_DATE_DDMMYYYY);
				break;
			case "datetime":
				$active_sheet->getStyle($cell_reference)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);
				break;
			case "time":
				$active_sheet->getStyle($cell_reference)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_DATE_TIME3);
				break;
		}
	}


	private function _get_formatting($element) {
		// Create a new format object
		$format = new ExcelFormats();
		$format_exists = false;

		// Find the element colour (if defined)
		$element_colour = $element->getAttribute('data-excel-colour');
		$element_font_colour = $element->getAttribute('data-excel-font-colour');

		// If there is a element colour
		if ($element_font_colour != '' && $element_font_colour != '') {
			// Note it within the sheet specification
			$format->font_colour = $element_font_colour;
			$format_exists = true;
		}

		// If there is a element colour
		if ($element_colour != '' && $element_colour != '') {
			// Note it within the sheet specification
			$format->background_colour = $element_colour;
			$format_exists = true;
		}

		// Find if the text should be struck through
		$element_strike_through = $element->getAttribute('data-excel-strike');

		if ($element_strike_through == true) {
			$format->strike_through = true;
			$format_exists = true;
		}

		// Find the data format
		$element_format = $element->getAttribute('data-excel-format');

		if ($element_format != '') {
			$format->data_format = $element_format;
			$format_exists = true;
		}

		$border_style = $element->getAttribute('data-border-style');
		if ( $border_style != '' ) {
			$format->use_border = $border_style;
			$format_exists = true;
		}

		// If there is a format
		if ($format_exists) {
			// Return it
			return $format;
		}
		else {
			// Return null
			return null;
		}
	}
}