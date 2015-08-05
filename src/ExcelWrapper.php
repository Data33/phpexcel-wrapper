<?php
    /**
    * @author Mattias Ottosson <datae33@gmail.com>
    * @link https://github.com/data33
    */
    namespace Data33\ExcelWrapper;

    use Data33\ExcelWrapper\Exceptions\ExcelException;
    use PHPExcel;
    use PHPExcel_Cell;
    use PHPExcel_IOFactory;

    class ExcelWrapper
    {
        private $rows, $title;

        function __construct(){
            $this->rows = [];
            $this->title = '';
        }

        /**
        * Add a new row to the document
        *
        * @param array $values An array with cell data for the row
        *
        * @return ExcelWrapper
        */
        public function addRow(array $values, $style = 'default', $line = null){
            //Take care of default input if we only want to supply a values array and an index
            if (is_integer($style)){
                $line = (int)$style;
                $style = 'default';
            }

            $row = [$values, $style];

            if($line === null){
                $this->rows[] = $row;
            }
            else{
                $this->rows[($line - 1)] = $row;
            }
            return $this;
        }

        /**
        * Save the document to a file path on the server
        *
        * @param string $filePath A path to where on the server file system the document should be saved
        *
        * @return ExcelWrapper
        */
        public function save($filePath){
            $objPHPExcel = new PHPExcel();
            $sheet = $objPHPExcel->setActiveSheetIndex(0);

            $rowNum = 1;

            $maxCol = 1;

            if (strlen($this->title)){
                $sheet->getStyle('A' . $rowNum)->applyFromArray(ExcelStyle::style('title'));
                $sheet->setCellValue('A' . $rowNum, $this->title);

                $rowNum += 2;
            }

            foreach($this->rows as $rowOffset => $row){
                list($rowCells, $style) = $row;
				
				//Update our maxcol so we can set the autosize later
                $maxCol = max($maxCol, count($rowCells)) - 1;
                $maxColString = PHPExcel_Cell::stringFromColumnIndex($maxCol);

                $offsetRowNum = ($rowNum + $rowOffset);

                //Set row style first
                $sheet->getStyle('A' . $offsetRowNum . ':' . $maxColString . $offsetRowNum)->applyFromArray(ExcelStyle::style($style));

                foreach($rowCells as $i => $cellValue){
                    $column = PHPExcel_Cell::stringFromColumnIndex($i);

                    if (is_array($cellValue) && count($cellValue) === 2){
                        list($cellValue, $cellStyle) = $cellValue;
                        $sheet->getStyle($column . $offsetRowNum)->applyFromArray(ExcelStyle::style($cellStyle));
                    }

                    $sheet->setCellValue($column . $offsetRowNum , $cellValue);
                }
            }
            //Loop through all used columns and set autosize to improve the looks of the document
            for($col = 'A'; $col <= $maxColString; $col++) {
                $sheet->getColumnDimension($col)
                    ->setAutoSize(true);
            }

            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
            $objWriter->save($filePath);
            $objPHPExcel->disconnectWorksheets();
            unset($objPHPExcel);

            return $this;
        }

        /**
        * Output the document to the browser and force download
        *
        * @param string $fileName The name of the resulting file to be downloaded by the browser
        *
        * @return ExcelWrapper
        */
        public function outputToBrowser($fileName){
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="' . $fileName);
            header('Cache-Control: max-age=0');

            //Use the save method on php's output stream
            $this->save('php://output');

            return $this;
        }

        /**
        * Set a document title/header
        * This will offset the line numbers when adding new lines to the document
        *
        * @param string $title The document title to be displayed in the file
        *
        * @return ExcelWrapper
        */
        public function setTitle($title){
            $this->title = $title;

            return $this;
        }
    }