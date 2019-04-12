<?php
require_once '../vendor/autoload.php';
$cal = new \om\IcalParser();
$results = $cal->parseFile(
	//'http://theodia.org/api/feed?distance=2&lang=fr&language=fr&latitude=46.5269159&longitude=6.901089500000012&rites=1&start=2019-02-27T23:00:00.000Z'
        'https://calendar.google.com/calendar/ical/emf.theodia%40gmail.com/private-e5c4e839e2928a23876e3b0d5ee542fc/basic.ics'
);

$events = $cal->getSortedEvents();

?>
<!DOCTYPE html>
<html lang="cs-CZ">
<head>
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	
	<title>Ical Parser example</title>
	
	<link href="//netdna.bootstrapcdn.com/bootstrap/3.0.3/css/bootstrap.min.css" rel="stylesheet">

</head>
<body>
<div class="container">
	<h1>Evenements</h1>
	
	<ul>
		<?php
		foreach ($events as $event) {
			$locationTab = explode(', ', $event['LOCATION']);
			$locationSize = sizeof($locationTab);
			$location = $locationTab[0] . ', ' . $locationTab[$locationSize-2];
			if ($event['isDateEvent']){
			    echo sprintf('  <li>le %s - Evènement du jour : %s - %s</li>' . PHP_EOL, $event['DTSTART']->format('d.m.Y'), $event['SUMMARY'], $event['DESCRIPTION']);
       
			} else {
				echo sprintf('	<li>le %s à %s - %s - %s - %s</li>' . PHP_EOL, $event['DTSTART']->format('d.m.Y'), $event['DTSTART']->format('H:i'), $location, $event['SUMMARY'], $event['DESCRIPTION']);
			}
		}
		?></ul>
</div>
</body>
</html>