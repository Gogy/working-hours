<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

if(!isset($_POST['submit'])) {
	echo "<h2>Working hours file generator</h2>";
	echo "<form action='' method='post'>
	Month <br />
	<select name='month'>";
		for($x=1;$x<=12;$x++) {
			echo "<option value='".$x."'>$x</option>";
		}
	echo "</select> <br />
	Year <br />
	<select name='year'>";
		for($y=2013;$y<=2033;$y++) {
			echo "<option value='".$y."'>$y</option>";
		}
	echo "</select><br /><br />
	<input type='submit' value='Yup, I am lazy' name='submit' />
	</form>";
}

else {

require_once 'PHPExcel.php';

function build_calendar($month,$year) {

$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Infobip Tools")
	 ->setLastModifiedBy("Infobip Tools")
	 ->setTitle("Working hours")
	 ->setSubject("Working hours")
	 ->setDescription("")
	 ->setKeywords("")
	 ->setCategory("Working Hours");

// Borders
$sharedStyle1 = new PHPExcel_Style();
$sharedStyle1->applyFromArray(
	array('borders' => array('bottom'=> array('style' => PHPExcel_Style_Border::BORDER_THIN),
				'right'=> array('style' => PHPExcel_Style_Border::BORDER_THIN),							)
		 ));

$objPHPExcel->getActiveSheet()->setSharedStyle($sharedStyle1, "A1:G1");

     $firstDayOfMonth = mktime(0,0,0,$month,1,$year);
     $numberDays = date('t',$firstDayOfMonth);
     $dateComponents = getdate($firstDayOfMonth);
     $monthName = $dateComponents['month'];

     $dayOfWeek = $dateComponents['wday'];

     $currentDay = 1;
     
     $month = str_pad($month, 2, "0", STR_PAD_LEFT);
  
     while ($currentDay <= $numberDays) {
          
          $currentDayRel = str_pad($currentDay, 2, "0", STR_PAD_LEFT);
          
          $date = "$year-$month-$currentDayRel";

		$timestamp = strtotime($date);

		$day = date('l', $timestamp);

          $currentDay++;
          $dayOfWeek++;

// Background color and values for first row
$objPHPExcel->getActiveSheet()->setSharedStyle($sharedStyle1, "A".$currentDay.":G".$currentDay);

$objPHPExcel->getActiveSheet()->setCellValue('A1', 'DATE')
				->setCellValue('B1', 'DAY')
				->setCellValue('C1', 'SHIFT')
				->setCellValue('D1', 'HOURS')
				->setCellValue('E1', 'NIGHT SHIFTS')
				->setCellValue('F1', 'SUNDAY')
				->setCellValue('G1', 'HOLIDAY');

// Legend
$objPHPExcel->getActiveSheet()->setCellValue('J1', 'Legend')
				->setCellValue('J2', 'FD')
				->setCellValue('J3', 'PV')
				->setCellValue('J4', 'SL')
				->setCellValue('K2', 'Free day')
				->setCellValue('K3', 'Paid vacation')
				->setCellValue('K4', 'Sick leave');


// Background color for sundays and saturdays
if($day == "Sunday") {
	$objPHPExcel->getActiveSheet()->getStyle('B'.$currentDay.':G'.$currentDay)->applyFromArray(
		array('fill'=> array('type'=> PHPExcel_Style_Fill::FILL_SOLID,'color'=> array('argb' => 'D67676'))));
} else if($day == "Saturday") {
	$objPHPExcel->getActiveSheet()->getStyle('B'.$currentDay.':G'.$currentDay)->applyFromArray(
		array('fill'=> array('type'=> PHPExcel_Style_Fill::FILL_SOLID,'color'=> array('argb' => '94AFD6'))));
}


// Write dates and days (first two columns)
$objPHPExcel->getActiveSheet()->setCellValue('A'.$currentDay, $date)
			      ->setCellValue('B'.$currentDay, $day);
}

// Write sums of working hours and but background colors
$objPHPExcel->getActiveSheet()->setCellValue('B'.($currentDay+1), 'Night shifts total:');
$objPHPExcel->getActiveSheet()->setCellValue('B'.($currentDay+2), 'Sundays total:');
$objPHPExcel->getActiveSheet()->setCellValue('B'.($currentDay+3), 'Holidays total:');
$objPHPExcel->getActiveSheet()->setCellValue('B'.($currentDay+4), 'Total:');
$objPHPExcel->getActiveSheet()->setCellValue('E'.($currentDay+1), '=SUM(E2:E'.($currentDay).')');
$objPHPExcel->getActiveSheet()->setCellValue('F'.($currentDay+2), '=SUM(F2:F'.($currentDay).')');
$objPHPExcel->getActiveSheet()->setCellValue('G'.($currentDay+3), '=SUM(G2:G'.($currentDay).')');
$objPHPExcel->getActiveSheet()->setCellValue('D'.($currentDay+4), '=SUM(D2:D'.($currentDay).')');

$objPHPExcel->getActiveSheet()->getStyle('A1:G1')->applyFromArray(
	array('fill'=> array('type'=> PHPExcel_Style_Fill::FILL_SOLID,'color'=> array('argb' => '3E61B5'))));

$objPHPExcel->getActiveSheet()->getStyle('A2:A'.$currentDay)->applyFromArray(
			array('fill'=> array('type'=> PHPExcel_Style_Fill::FILL_SOLID,'color'=> array('argb' => '5F8FAD'))));

$objPHPExcel->getActiveSheet()->getStyle('A'.($currentDay+1).':G'.($currentDay+1))->applyFromArray(
			array('fill'=> array('type'=> PHPExcel_Style_Fill::FILL_SOLID,'color'=> array('argb' => 'F0C20A'))));
$objPHPExcel->getActiveSheet()->getStyle('A'.($currentDay+2).':G'.($currentDay+2))->applyFromArray(
			array('fill'=> array('type'=> PHPExcel_Style_Fill::FILL_SOLID,'color'=> array('argb' => 'A6D1FF'))));
$objPHPExcel->getActiveSheet()->getStyle('A'.($currentDay+3).':G'.($currentDay+3))->applyFromArray(
			array('fill'=> array('type'=> PHPExcel_Style_Fill::FILL_SOLID,'color'=> array('argb' => 'DEC1D5'))));
$objPHPExcel->getActiveSheet()->getStyle('A'.($currentDay+4).':D'.($currentDay+4))->applyFromArray(
			array('fill'=> array('type'=> PHPExcel_Style_Fill::FILL_SOLID,'color'=> array('argb' => '5F8FAD'))));


// Center text for everything and adjust column sizes
$objPHPExcel->getActiveSheet()->getStyle('A1:G'.($currentDay+4))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(6);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(12);


// Format date columns and set numbers to decimals
PHPExcel_Cell::setValueBinder( new PHPExcel_Cell_AdvancedValueBinder() );

$objPHPExcel->getActiveSheet()
           ->getStyle('A2:A'.$currentDay)
           ->getNumberFormat()
           ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DDMMYYYY2);

$objPHPExcel->getActiveSheet()
           ->getStyle('D2:G'.($currentDay+4))
           ->getNumberFormat()
           ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);


// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle('Working hours');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);

// Type of file
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');


// We'll be outputting an excel file
header('Content-type: application/vnd.ms-excel');

// It will be called file.xls
header('Content-Disposition: attachment; filename="'.$month.'-'.$year.'.xlsx"');

// Write file to the browser
$objWriter->save('php://output');

$objPHPExcel->getActiveSheet()->setCellValue('A1', 'DATE');

}


	
     $dateComponents = getdate();

     $month = $_POST['month'];		     
     $year = $_POST['year'];

     build_calendar($month,$year);


}


                    


?>
