<?php
/**
 * Created by PhpStorm.
 * User: morattelp
 * Date: 21.02.2019
 * Time: 13:43
 */

include_once 'resources/event_data.php';
require_once '../vendor/autoload.php';

$phpWord = new \PhpOffice\PhpWord\PhpWord();
//$templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor('resources/test_template3.docx');
																				//link format  'https://drive.google.com/uc?id=[FILE_ID]&export=download'
$templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor('https://drive.google.com/uc?id=1T5U3N07WYq7FdR3_bvwzWuuXSw5qeafH&export=download');
$variables = $templateProcessor->getVariables();
$groupDays = (in_array('beginRow2', $variables));

$parser = new HTMLtoOpenXML\Parser();

$templateProcessor->setImageValue('logo', 'http://www.upsaintjoseph.ch/typo3temp/_processed_/6/a/csm_logo-header_95aacd674a.png'); //'resources/logo_emf_phasePro.jpg');

//Test insertion HTML
\PhpOffice\PhpWord\Settings::setOutputEscapingEnabled(false);
$htmlString = '<ul>
				<li>test</li>
				<li>test</li>
				<li>test</li>
			</ul>';
$htmlString = $parser->fromHTML($htmlString);
$templateProcessor->setValue('testHTML', $htmlString);
\PhpOffice\PhpWord\Settings::setOutputEscapingEnabled(true);

if ($groupDays) {
	
	/*
    * group events by dates
    */
	$days = [];
	$oldDate = null;
	$i = 0;
	foreach ($events as $event){
		if ($oldDate !== null){
			if ($oldDate == $event["date"]){
				$days[$i][] = $event;
			} else {
				$oldDate = $event["date"];
				$i++;
				$day = [];
				$day[]= $event;
				$days[$i] = $day;
			}
		} else {
			$oldDate = $event["date"];
			$day = [];
			$day[]= $event;
			$days[$i] = $day;
		}
	}
	
	$templateProcessor->cloneRow('beginRow1', sizeof($days));
	for ($i = 0; $i < sizeof($days); $i++) {
		$templateProcessor->setValue('date#' . ($i + 1), $days[$i][0]['date']);
		$templateProcessor->cloneRow('beginRow2#' . ($i + 1), sizeof($days[$i]));
		
		for ($j = 0; $j < sizeof($days[$i]); $j++) {
			$templateProcessor->setValue('location#' . ($i + 1) . '#' . ($j + 1), $days[$i][$j]['location']);
			$templateProcessor->setValue('time#' . ($i + 1) . '#' . ($j + 1), $days[$i][$j]['time']);
			$templateProcessor->setValue('summary#' . ($i + 1) . '#' . ($j + 1), $days[$i][$j]['summary']);
			$templateProcessor->setValue('description#' . ($i + 1) . '#' . ($j + 1), $days[$i][$j]['description'] === '' ? '' : '  : ' . $days[$i][$j]['description']);
			
		}
	}
} else {
	$templateProcessor->cloneRow('beginRow1', sizeof($events));
	for ($i = 0; $i < sizeof($events); $i++) {
		$templateProcessor->setValue('date#' . ($i + 1), $events[$i]['date']);
		$templateProcessor->setValue('location#' . ($i + 1), $events[$i]['location']);
		$templateProcessor->setValue('time#' . ($i + 1), $events[$i]['time']);
		$templateProcessor->setValue('summary#' . ($i + 1), $events[$i]['summary']);
		$templateProcessor->setValue('description#' . ($i + 1), $events[$i]['description']);
	}
}

$templateProcessor->setValue('dateEvent#1', 'Evenement date 1');

foreach ($templateProcessor->getVariables() as $variable) {
	if (substr($variable, 0, 8) === 'beginRow' || substr($variable, 0, 9) === 'dateEvent'){
		$templateProcessor->setValue($variable, '');
	}
}

//Output result
$fileName = 'test_template_phpWord.docx';
$templateProcessor->saveAs('results/test_template_phpWord.docx');

$contentType = 'Content-type: application/vnd.openxmlformats-officedocument.wordprocessingml.document;';
header ( "Expires: Mon, 1 Apr 1974 05:00:00 GMT" );
header ( "Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT" );
header ( "Cache-Control: no-cache, must-revalidate" );
header ( "Pragma: no-cache" );
header ( $contentType );
header ( "Content-Disposition: attachment; filename=" . $fileName);
readfile('results/' . $fileName);