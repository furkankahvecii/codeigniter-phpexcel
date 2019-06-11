<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class PHPExcel_example extends CI_Controller {
	public function index()
	{
        // Read an Excel File
        $tmpfname = "example.xls";
        $excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
        $objPHPExcel = $excelReader->load($tmpfname);
        
        // Set document properties
        $objPHPExcel->getProperties()->setCreator("Furkan Kahveci")
							 ->setLastModifiedBy("Furkan Kahveci")
							 ->setTitle("Office 2007 XLS Test Document")
							 ->setSubject("Office 2007 XLS Test Document")
							 ->setDescription("Description for Test Document")
							 ->setKeywords("phpexcel office codeigniter php")
							 ->setCategory("Test result file");

        // Create a first sheet
        $objPHPExcel->setActiveSheetIndex(0);
        $objPHPExcel->getActiveSheet()->setCellValue('A1', "Furkan");
        $objPHPExcel->getActiveSheet()->setCellValue('B1', "Kahveci");
        $objPHPExcel->getActiveSheet()->setCellValue('C1', "KLU");
        $objPHPExcel->getActiveSheet()->setCellValue('D1', "Software Engineering");
        $objPHPExcel->getActiveSheet()->setCellValue('E1', "11.06.2019");

        // Hide F and G column
        $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setVisible(false);
        $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setVisible(false);

        // Set auto size
        $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);

        // Add data
        for ($i = 10; $i <= 50; $i++) 
        {
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, "Name $i")
                                        ->setCellValue('B' . $i, "Surname $i")
                                        ->setCellValue('C' . $i, "University $i")
                                        ->setCellValue('D' . $i, "Department $i")
                                        ->setCellValue('E' . $i, "Date");
        }

        // Set Font Color, Font Style and Font Alignment
        $stil=array(
            'borders' => array(
              'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('rgb' => '000000')
              )
            ),
            'alignment' => array(
              'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            )
        );
        $objPHPExcel->getActiveSheet()->getStyle('A3:E3')->applyFromArray($stil);

        // Merge Cells
        $objPHPExcel->getActiveSheet()->mergeCells('A5:E5');
        $objPHPExcel->getActiveSheet()->setCellValue('A5', "MERGED CELL");
        $objPHPExcel->getActiveSheet()->getStyle('A5:E5')->applyFromArray($stil);
        
        // Save Excel xls File
        $filename="filename.xls";
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        ob_end_clean();
        header('Content-type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename='.$filename);
        $objWriter->save('php://output');
	}
}
