<?php
namespace Trunk\ExcelLibrary\Excel;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once( __dir__ . '/ExcelSheet.php' );

class Excel extends \Trunk\Wibbler\Modules\base {

	var $author = "Wibbler";
	var $description = "Automatically generated from Wibbler";

	private $company = '';
	private $title = "Wibbler Report";

	/**
	 * @var ExcelSheet[]
	 */
	var $sheet;

	var $sheetCount = 1;

	public function __construct() {
		/*
		 *		memory set at 1GB to cater for for very large memory requirements
		 */
		ini_set( 'memory_limit', '1024M' );

		// prevent foreach error when there are no sheets
		$this->sheet = array();
	}

	#region Setters
	public function setAuthor( $author ) {
		$this->author = $author;
		return $this;
	}

	public function setCompany( $company ) {
		$this->company = $company;
		return $this;
	}

	public function setDescription( $description ) {
		$this->description = $description;
		return $this;
	}
	#endregion

	public function loadFromHTML( $html ) {
		$a = new \DOMDocument();
		$a->loadHTML( $html );

		foreach ( $a->getElementsByTagName( 'table' ) as $table_index => $table ) {
			if ( $table->hasAttribute( 'data-excel-sheet' ) ) {
				$this->sheet[ $table_index ] = new \Trunk\ExcelLibrary\Excel\ExcelSheet();
				$sheetName = $table->getAttribute( 'data-excel-sheet-title' );
				if ( empty( $sheetName ) ) {
					$sheetName = "Sheet" . $table_index;
				}
				$this->sheet[ $table_index ]->sheetTitle = preg_replace( "/[^ \\w]+/", " ", substr( $sheetName, 0, 31 ) );
				$this->sheet[ $table_index ]->sheetDescription = $table->getAttribute( 'data-excel-sheet-description' );

				// Get the thead's row object (only the first one though)
				$thead = $table->getElementsByTagName( 'thead' )->item( 0 );
				$this->sheet[ $table_index ]->num_header_rows = $thead->getElementsByTagName( 'tr' )->length;

				// Process the header rows
				$this->sheet[ $table_index ]->end_column_number = $this->_process_header( $table_index, $thead );

				// Get the tbody
				$tbody = $table->getElementsByTagName( 'tbody' )->item( 0 );
				// Process the body data into the array of data to show
				$this->_process_body( $table_index, $tbody );

				// Try to find the tfoot (often won't exist
				$tfoot = $table->getElementsByTagName( 'tfoot' )->item( 0 );
				// Process the footer data into the array of data to show
				$this->_process_footer( $table_index, $tfoot );
			}
		}
	}

	/**
	 * Process the data from the thead section of the table
	 * @param $table_index
	 * @param \DOMNode $thead
	 * @return int
	 */
	private function _process_header( $table_index, \DOMNode $thead ) {

		$num_data_columns = 0;
		$header_rows = [ ];

		foreach ( $thead->getElementsByTagName( 'tr' ) as $row_index => $tr ) {

			// Create an empty row of cells for this header row
			if ( !isset( $header_rows[ $row_index ] ) ) {
				$header_rows[ $row_index ] = [ ];
			}

			$current_col = 0;
			foreach ( $tr->getElementsByTagName( 'th' ) as $col_index => $th ) {

				// Find the number of columns to span
				$num_cols = $th->getAttribute( 'colspan' );
				$num_rows = $th->getAttribute( 'rowspan' );
				$text = $th->textContent;

				if ( empty( $num_cols ) )
					$num_cols = 1;
				if ( empty( $num_rows ) )
					$num_rows = 1;

				// If we have row spanning
				/*if ( $num_rows > 1 ) {
					// Note on the second row that this cell has already been spanned
					$header_rows[ 1 ][ $current_col ] = "Span";
				}

				while ( $row_index == 1 && isset( $header_rows[ 1 ][ $current_col ] ) && $header_rows[ 1 ][ $current_col ] == "Span" ) {
					$current_col++;
				}*/

				while ( $row_index > 0 && isset( $header_rows[ $row_index ][ $current_col ] ) && $header_rows[ $row_index ][ $current_col ] == "Span" ) {
					$current_col++;
				}

				for ( $i = $row_index; $i < ( $row_index + $num_rows ); $i++ ) {
					if ( $num_cols ) {
						for ( $j = $current_col; $j < ( $current_col + $num_cols ); $j++ ) {
							$header_rows[ $i ][ $j ] = "Span";
						}
					} else {
						$header_rows[ $i ][ $current_col ] = "Span";
					}
				}

				$header_cell = new ExcelHeader( $text, $current_col, $num_cols, $num_rows );
				$header_rows[ $row_index ][ $current_col ] = $header_cell;

				$current_col += $header_cell->col_span;
			}

			// If we're processing the first row
			if ( $row_index == 0 ) {
				// Note the total number of data columns
				$num_data_columns += $current_col;
			}
		}

		$this->sheet[ $table_index ]->header_rows = $header_rows;
		return $num_data_columns;
	}

	/**
	 * Processes the data from the tbody section of the table
	 * @param int $table_index
	 * @param \DOMNode $tbody
	 */
	private function _process_body( $table_index, \DOMNode $tbody ) {

		$xl = [ ];

		foreach ( $tbody->getElementsByTagName( 'tr' ) as $row_index => $tr ) {
			$this->sheet[ $table_index ]->add_row_format( $row_index, $tr );

			foreach ( $tr->getElementsByTagName( 'td' ) as $col_index => $td ) {
				$formula = $td->getAttribute( 'data-formula' );
				if ( $formula != "" ) {
					$xl[ $row_index ][ $col_index ] = "" . $formula;
				}
				else {
					$xl[ $row_index ][ $col_index ] = "" . $td->textContent;
				}

				$this->sheet[ $table_index ]->add_cell_format( $row_index . '.' . $col_index, $td );
			}
		}
		$this->sheet[ $table_index ]->data = $xl;
	}

	/**
	 * Process the data from the tfoot section of the table
	 * @param int $table_index
	 * @param \DOMNode $tfoot
	 */
	private function _process_footer( $table_index, $tfoot ) {

		// If there is no footer
		if ( $tfoot == null ) {
			return;
		}

		// Get the footer row (only one)
		$footer_row = $tfoot->getElementsByTagName( 'tr' )->item( 0 );

		// If there is now footer row
		if ( $footer_row == null ) {
			return;
		}

		$num_rows = count( $this->sheet[ $table_index ]->data );
		$footer_format = new ExcelFormats();
		$footer_format->background_colour = 'E0E0FF';
		$footer_format->bold = true;
		$this->sheet[ $table_index ]->set_footer_format( $num_rows, $footer_format );

		foreach ( $footer_row->getElementsByTagName( 'td' ) as $col_index => $tf ) {
			if ( !empty( $tf->textContent ) ) {
				$this->sheet[ $table_index ]->data[ $num_rows ][ $col_index ] = "" . $tf->textContent;
			}
		}
	}

	public function create( $output_to = null, $report_name = 'Report' ) {
		$this->title = $report_name;

		while ( ob_get_level() > 0 ) {
			ob_end_clean();
		}

		// Create new PHPExcel object
		$objPHPExcel = new Spreadsheet();
		// Set properties
		$objPHPExcel->getProperties()->setCreator( $this->author );
		$objPHPExcel->getProperties()->setLastModifiedBy( $this->author );
		$objPHPExcel->getProperties()->setTitle( $this->title );
		$objPHPExcel->getProperties()->setSubject( $this->title );
		$objPHPExcel->getProperties()->setDescription( $this->description );
		$objPHPExcel->getProperties()->setCompany( $this->company );

		$current_sheet_index = 0;
		foreach ( $this->sheet as $thisSheet ) {
			//Add new sheet
			if ( $current_sheet_index > 0 )
				$objPHPExcel->createSheet();

			$objPHPExcel->setActiveSheetIndex( $current_sheet_index );
			$active_sheet = $objPHPExcel->getActiveSheet();

			$endColumnLetter = $thisSheet->get_end_column_letter();

			if ( $thisSheet->sheetTitle == "" )
				$thisSheet->sheetTitle = $this->title;

			//Add head rows
			// Merge the cells
			$active_sheet->mergeCells( "A1:" . $endColumnLetter . "1" );
			$active_sheet->mergeCells( "A2:" . $endColumnLetter . "2" );
			$active_sheet->mergeCells( "A3:" . $endColumnLetter . "3" );

			// Set the header cell contents
			$active_sheet->SetCellValue( "A1", $this->title );
			$active_sheet->SetCellValue( "A2", $thisSheet->sheetTitle );
			$active_sheet->SetCellValue( "A3", $thisSheet->sheetDescription );

			// Set the background styling for the header cells
			$active_sheet->getStyle( "A1:A4" )->getFill()->setFillType( Fill::FILL_SOLID );
			$active_sheet->getStyle( "A1:A4" )->getFill()->getStartColor()->setARGB( Color::COLOR_WHITE );

			// Style the header cells
			$active_sheet->getStyle( 'A1' )->getFont()->setSize( 20 );
			$active_sheet->getStyle( 'A1' )->getFont()->setBold( true );
			$active_sheet->getStyle( 'A2' )->getFont()->setSize( 14 );
			$active_sheet->getStyle( 'A2' )->getFont()->setBold( true );

			// Define which row the table headers are
			$headerRowNum = 5;
			$current_row_num = $headerRowNum;

			foreach ( $thisSheet->header_rows as $header_row ) {

				$column_index = 0;
				foreach ( $header_row as $header_cell ) {

					if ( $header_cell == "Span" )
						continue;

					$columnLetter = ExcelSheet::get_letter_from_number( $header_cell->start_col );

					//Set column header
					$active_sheet->SetCellValue( $columnLetter . $current_row_num, $header_cell->title );

					// If the column is set to span multiple cells
					if ( $header_cell->col_span > 1 || $header_cell->row_span > 1 ) {
						$merge_cells = $columnLetter . $current_row_num . ":" . ExcelSheet::get_letter_from_number( $header_cell->start_col + $header_cell->col_span - 1 ) . ( $current_row_num + $header_cell->row_span - 1 );
						// Set the spanning
						$active_sheet->mergeCells( $merge_cells );
						$active_sheet->getStyle( $merge_cells )->getAlignment()->setHorizontal(  Alignment::HORIZONTAL_CENTER );
					}

					$column_index += $header_cell->col_span;
				}

				$current_row_num++;
			}

			// Set the formatting for all of the header cells in one call
			$header_range = "A" . $headerRowNum . ":" . $thisSheet->end_column_letter . ( $headerRowNum + $thisSheet->num_header_rows - 1 );
			$active_sheet->getStyle( $header_range )->getFill()->setFillType( Fill::FILL_SOLID );
			$active_sheet->getStyle( $header_range )->getFill()->getStartColor()->setARGB( 'FFE0E0FF' );
			$active_sheet->getStyle( $header_range )->getFont()->setBold( true );

			// Loop through each column outputting the data
			for ( $col = 0; $col < $thisSheet->end_column_number; $col++ ) {
				// Reset the first row to use
				$rowNumber = $headerRowNum + $thisSheet->num_header_rows;
				// Output the column's data
				$rowNumber = $this->_set_column_data( $thisSheet, $active_sheet, $col, $rowNumber );
			}

			$active_sheet->getStyle( "A" . ( $rowNumber ) . ":" . $endColumnLetter . ( $rowNumber ) )->getFill()->setFillType( Fill::FILL_SOLID );
			$active_sheet->getStyle( "A" . ( $rowNumber ) . ":" . $endColumnLetter . ( $rowNumber ) )->getFill()->getStartColor()->setARGB( 'FFE0E0FF' );

			$active_sheet->mergeCells( "A" . ( $rowNumber + 2 ) . ":" . $endColumnLetter . ( $rowNumber + 2 ) );
			$active_sheet->SetCellValue( "A" . ( $rowNumber + 2 ), "Created " . date( 'd/m/Y H:i' ) );

			// Name sheet
			$active_sheet->setTitle( ( $thisSheet->tabTitle == null ? $thisSheet->sheetTitle : $thisSheet->tabTitle ) );

			//Set sheet priting properties
			$active_sheet->getHeaderFooter()->setOddHeader( '&L&B' . $thisSheet->sheetTitle . '&RPrinted on &D' );
			$active_sheet->getHeaderFooter()->setOddFooter( '&L&B' . $objPHPExcel->getProperties()->getTitle() . '&RPage &P of &N' );
			$active_sheet->getPageSetup()->setFitToPage( true );
			$active_sheet->getPageSetup()->setFitToWidth( 1 );
			$active_sheet->getPageSetup()->setFitToHeight( 0 );

			$current_sheet_index++;
		}

		//Set active sheet to first one
		$objPHPExcel->setActiveSheetIndex( 0 );

		// Create writer object
		$writer = new Xlsx( $objPHPExcel );

		if ( $output_to == null ) {
			//Set headers so output is file
			header( 'Content-type: application/ms-excel' );
			$now = new \DateTime();
			header( 'Content-Disposition: attachment; filename="' . substr( $report_name, 0, 31 ) . ' ' . $now->format( 'Y-m-d H:i' ) . '.xlsx"' );
			flush();

			//Output document
			$writer->save( 'php://output' );
		}
		else {

			// Output to the given file
			$writer->save( $output_to );
		}

		unset( $this->objPHPExcel );
		unset( $writer );
#		$this->objPHPExcel = Null ;
#		$writer = Null ;
	}

	private function _set_column_data( $thisSheet, $active_excel_sheet, $column_index, $start_row ) {

		$rowNumber = $start_row;
		$columnDataKey = $column_index;
		$columnLetter = ExcelSheet::get_letter_from_number( $column_index );

		// If we have some data
		if ( !empty( $thisSheet->data ) ) {
			// Iterate over the data
			foreach ( $thisSheet->data as $key => $dataObject ) {
				$thisSheet->set_row_format( $active_excel_sheet, $rowNumber, $key );

				// If a specific format has been requested
				if ( !empty( $dataObject[ $columnDataKey ] ) && $thisSheet->formatMap != null && isset( $thisSheet->formatMap[ $column_index ] ) ) {
					switch ( $thisSheet->formatMap[ $column_index ] ) {
						case 'Date':
							// Add 3600 to make sure the date is still today - covers BST
							$timestamp = strtotime( $dataObject[ $columnDataKey ] ) + 3600;
							if ( $timestamp % ( 24 * 3600 ) == 3600 )
								$timestamp = $timestamp - 3600;
							$active_excel_sheet->SetCellValue( $columnLetter . $rowNumber, \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel( $timestamp ) );
							$active_excel_sheet->getStyle( $columnLetter . $rowNumber )->getNumberFormat()->setFormatCode( 'd mmm yyyy' );
							break;
						case 'Time':
							//Enter data for this row
							$active_excel_sheet->SetCellValue( $columnLetter . $rowNumber, $this->time_to_excel( $dataObject[ $columnDataKey ] ) );
							$active_excel_sheet->getStyle( $columnLetter . $rowNumber )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_DATE_TIME3 );
							break;
						case 'Text':
							$active_excel_sheet->SetCellValueExplicit( $columnLetter . $rowNumber, $dataObject[ $columnDataKey ], DataType::TYPE_STRING );
//									$active_sheet->SetCellValue($columnLetter . $rowNumber, $dataObject[ $columnDataKey ]);
//									$active_sheet->getStyle($columnLetter . $rowNumber)->getNumberFormat()->setFormatCode('0');
							break;
						default:
							//Enter data for this row
							$active_excel_sheet->SetCellValue( $columnLetter . $rowNumber, $dataObject[ $columnDataKey ] );
							break;
					}
				}
				else {
					//Enter data for this row
					$active_excel_sheet->SetCellValue( $columnLetter . $rowNumber, isset( $dataObject[ $columnDataKey ] ) ? $dataObject[ $columnDataKey ] : "" );
				}

				$thisSheet->set_cell_format( $active_excel_sheet, $columnLetter . $rowNumber, $key . '.' . $column_index );

				$rowNumber++;
			}
		}

		$active_excel_sheet->getColumnDimension( $columnLetter )->setAutoSize( true );

		return $rowNumber;
	}

	function time_to_excel( $time = '00:00:00' ) {
		list( $hours, $mins, $secs ) = explode( ':', $time );
		$seconds = ( $hours * 3600 ) + ( $mins * 60 ) + $secs;

		$day_seconds = 24 * 60 * 60;

		return $seconds / $day_seconds;
	}
}

/* End of file Excel.php */