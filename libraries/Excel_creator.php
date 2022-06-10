<?php defined('BASEPATH') or exit('No direct script access allowed');

use \PhpOffice\PhpSpreadsheet\Reader\Xls as Excel;
use \PhpOffice\PhpSpreadsheet\Worksheet\Protection;
use \PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;

use CI\Models\Model;
use CI\Models\Variant;

class Excel_creator{

	protected $ci;
	protected $str_html_format;
	protected $generated_file_name;
	protected $spreadsheet;
	protected $worksheet;
	protected $filename;
	protected $errors = [];
	protected $unprotected_cols = [];
	protected $file_path;
	protected $allowed_extensions = ['xls', 'xlsx'];

	function __construct(){
		$this->ci =& get_instance();
	}

	function from_html($str=NULL){
		if($str){
			$reader = new \PhpOffice\PhpSpreadsheet\Reader\Html();
			$this->spreadsheet = $reader->loadFromString($str);
			// dd($this->spreadsheet);
		}

		return $this;
	}

	/**
	 * Initialize spreadsheet it new spreadsheet object
	 * @return [type] [description]
	 */
	function spreadsheet(){
		if(!$this->spreadsheet) $this->spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet;
		return $this->spreadsheet;
	}

	function worksheet($to_be_activate_sheet=NULL){
		if(!$this->spreadsheet) $this->spreadsheet();

		if($to_be_activate_sheet) $this->worksheet = $this->spreadsheet->getSheet($to_be_activate_sheet);
		else $this->worksheet = $this->spreadsheet->getActiveSheet();

		return $this;
	}

	function is_protected($is_protected=TRUE, $password=NULL){
		if($is_protected){

			$protection = $this->worksheet->getProtection();
			$protection->setSheet(true);
			$protection->setPassword($password);
			$protection->setSheet(true);
			$protection->setSort(true);
			$protection->setInsertRows(true);
			$protection->setFormatCells(true);	
		}
	}

	function add_sheet($options=[]){
		$options = (object)$options;

		if(!$this->has_error()){
			$this->worksheet = $this->spreadsheet->createSheet();
			$this->worksheet->setTitle($options->title);
		}

		return $this;
	}

	function filename($filename=NULL){
		if($filename AND !$this->filename) $this->filename = $filename;
		return $this->filename;
	}

	function file_path($path=NULL){
		if($path) $this->file_path = $path;
		return $this;
	}

	function read(){

		if(!$this->file_path OR !file_exists($this->file_path)) $this->errors[] = 'Read: Please provide first the file path to be read.';
		if(!$this->is_valid_file_extension()) $this->errors[] = sprintf('Read: File extension is not allowed to be read. Allowed file are %1$s', implode(', ', $this->allowed_extensions));

		if(!$this->has_error()){
			$file_extension = $this->get_file_extension();

			if($file_extension === 'xls') $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls;
			elseif($file_extension === 'xlsx') $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx;

			$reader->setReadDataOnly(true);

			$this->spreadsheet = $reader->load($this->file_path);
		}

		return;
	}

	function cell($cell_or_col=NULL, $value=NULL, $options=[]){
		$options = (object)$options;


		if(!$cell_or_col) $this->errors[] = 'Cell: target cell and cell value is required.';

		if(!$this->has_error()){
			$worksheet = $this->worksheet;

			if(is_array($value)){
				$worksheet->fromArray($value, NULL, $cell_or_col);
			}
			else{
				if(preg_match('/[A-Z][0-9]+/', $cell_or_col)) $current_cell = $worksheet->getCell($cell_or_col);

				if($value AND $current_cell){
					$current_cell->setValue($value);
				}

				if(property_exists($options, 'protection')){
					// TODO Add protection into the cell, if the given cell is only column get the highest row and and protection to it
					$this->set_cell_protection($cell_or_col, $options->protection);
				}

				if(property_exists($options, 'style')){

					$this->set_cell_style($cell_or_col, $options->style);
				}
				
			}
		}

		return $this;
	}

	function set_cell_protection($cell_or_col=NULL, $protection_config=[]){
		$protection_config = (object)$protection_config;

		if(!$this->worksheet) $this->errors[] = 'Cell protection: unable to set protection on cell or column worksheet variable is null';

		if(!$this->has_error()){
			$worksheet = $this->worksheet;

			$highest_row = $worksheet->getHighestRow();
			$highest_column = $worksheet->getHighestColumn();

			if(preg_match('/[A-Z][0-9]+/', $cell_or_col) AND !property_exists($protection_config, 'start_row')){

				if(property_exists($protection_config, 'is_protected') AND !$protection_config->is_protected){
					$worksheet->getStyle($cell_or_col)->getProtection()
		    				->setLocked(\PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_UNPROTECTED);

				}
			}
			else{
				$start_row = (property_exists($protection_config, 'start_row')) ? $protection_config->start_row : 1;

				for($row=$start_row;$row<=$highest_row;$row++) {
					if(property_exists($protection_config, 'is_protected') AND !$protection_config->is_protected){
						$worksheet->getStyle(sprintf('%1$s%2$s', $cell_or_col, $row))->getProtection()
		    				->setLocked(\PhpOffice\PhpSpreadsheet\Style\Protection::PROTECTION_UNPROTECTED);
					}
				}
			}

		}
		return;
	}

	function set_cell_style($cell_or_col=NULL, $style_config=[]){
		$style_config = (object)$style_config;


		if(!$this->worksheet) $this->errors[] = 'Cell style: unable to set style on cell or column worksheet variable is null.';

		if(!$this->has_error()){
			$worksheet = $this->worksheet;

			if(property_exists($style_config, 'is_auto_size') AND is_bool($style_config->is_auto_size)){
				$column = str_replace('/[0-1]/', '', $cell_or_col);
				$worksheet->getColumnDimension($column)->setAutoSize(true);
			}

		}
	}

	function data_lists($column=NULL, $rows=NULL, $config=[]){
		//validate
		$config = (object)$config;
		if(!$column AND !$rows) $this->errors[] = 'Dropdown cell: columns and rows are required.';
		// if(!is_array($options)) $this->errors[] = 'Dropdown cell: option must be array.';
		if(!preg_match('/(\d+)-(\d+)|(\d+)/', $rows)) $this->errors[] = 'Dropdown cell: rows must in format of (start_index-end_index) OR Cell row only';

		if(!$this->has_error()){
			$worksheet = $this->worksheet;

			$rows = explode('-', $rows);

			if(count($rows) === 1){

			}
			else{
				$row_start = $rows[0];
				$row_end = $rows[1];

				for($counter=$row_start; $counter<=$row_end; $counter++){
			    	$validation = new DataValidation;
				    $validation->setType(DataValidation::TYPE_LIST);
				    $validation->setErrorStyle( DataValidation::STYLE_INFORMATION);
				    $validation->setAllowBlank(false);
					$validation->setShowInputMessage(true);
					$validation->setShowErrorMessage(true);
					$validation->setShowDropDown(true);
					$validation->setErrorTitle('Input error');
					$validation->setError('Value is not in list.');
					$validation->setPromptTitle('Pick from list');
					$validation->setPrompt('Please pick a value from the drop-down list.');
					$validation->setFormula1($config->formula);
					$cell = sprintf('%1$s%2$s', $column, $counter);
					$validation = $worksheet->getCell($cell)->setDataValidation($validation);
		    		// $validation->setFormula1(sprintf('"%1$s"', $options));
				}
			}
		}
	}

	function download(){

		if(!$this->spreadsheet) $this->errors[] = 'Download: unable to download spreadsheet object is null';

		if(!$this->has_error()){
		    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
			header(sprintf('Content-Disposition: attachment; filename="%1$s.xlsx"', $this->generate_file_name()));
			$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, 'Xlsx');
			// $writer->save(sprintf('%1$s.xlsx',  $this->generate_file_name())); 

			$writer->save('php://output');
			exit;
		}
	}

	function errors(){
		return $this->errors;
	}

	function has_error(){
		return count($this->errors) ? TRUE : FALSE;
	}

	function generate_file_name(){
		if(!$this->generated_file_name){
			$date_today = date('m-d-Y', now());

			if($this->filename) $this->generated_file_name = sprintf('%1$s-%2$s', $this->filename, $date_today);
			else $this->generated_file_name = $date_today;
		}
		return $this->generated_file_name;
	}

	protected function get_file_extension(){
		if($this->file_path){
			$path_part = pathinfo($this->file_path);
			return $path_part['extension'];
		}

		return NULL;
	}

	protected function is_valid_file_extension(){
		if($extension = $this->get_file_extension()){
			if(in_array($extension, $this->allowed_extensions)) return TRUE;
			return FALSE;
		}

		return FALSE;
	}
}