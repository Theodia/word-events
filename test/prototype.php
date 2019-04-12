<?php
require_once '../vendor/autoload.php';

$phpWord = new \PhpOffice\PhpWord\PhpWord();
$icalParser = new \om\IcalParser();
$HTMLparser = new HTMLtoOpenXML\Parser();

$results = $cal->parseFile(
	'https://calendar.google.com/calendar/ical/emf.theodia%40gmail.com/private-e5c4e839e2928a23876e3b0d5ee542fc/basic.ics'
);

$evenements = $cal->getSortedEvents();
$days = [];
$oldDate = null;
$i = 0;
foreach ($evenements as $event) {
	if ($oldDate !== null){
	
	} else {
		$oldDate = $event['DTSTART'];
		$i++;
		$day = [];
		$day[]=$event;
		$days[$i]= $day;
	}
}