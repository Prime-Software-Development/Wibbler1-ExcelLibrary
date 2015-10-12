<?php
namespace Trunk\ExcelLibrary\Excel;

class ExcelHeader {
	var $title = "";
	var $col_span = 1;
	var $row_span = 1;
	var $start_col = 0;

	public function __construct( $title, $start_column, $colspan = 1, $rowspan = 1 ) {
		$this->title = $title;
		$this->start_col = $start_column;
		$this->col_span = $colspan > 0 ? $colspan : 1;
		$this->row_span = $rowspan > 0 ? $rowspan : 1;
	}
}
