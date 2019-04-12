<?php

use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;

require_once '../../vendor/autoload.php';
include_once 'wrk/EventReader.php';
include_once 'wrk/DocGenerator.php';


const PARAM_DATEFORMAT = "dateFormat";
const PARAM_TEMPLATEURL = "templateURL";
const PARAM_MAINCALURL = "calendarMainURL";
const PARAM_OTHERCALURL = "calendarOtherURL";
const FILENAME = "fiche_dominicale.docx";


if (!empty($_POST[PARAM_DATEFORMAT]) && !empty($_POST[PARAM_TEMPLATEURL]) && !empty($_POST[PARAM_MAINCALURL])) {
	
	$errorMessage = "HTTP Error 500 : Internal Server Error";
	$gotException = false;
	if (file_exists('../results/' . FILENAME)){
		unlink('../results/' . FILENAME);
	}
	
	$eventReader = new EventReader($_POST[PARAM_DATEFORMAT]);
	$docGenerator = new DocGenerator($_POST[PARAM_TEMPLATEURL]);
	
	$organisationInfo = [
		't.organisation.logo' => $_POST['logoURL'],
		't.organisation.name' => $_POST['organisationName'],
		't.organisation.location' => $_POST['organisationLocation'],
		't.organisation.address' => $_POST['organisationAddress'],
		't.organisation.phonenumber' => $_POST['organisationPhoneNbr'],
	];
	
	try {
		$mainEvents = $eventReader->getEvents($_POST[PARAM_MAINCALURL]);
		
		if (!empty($_POST[PARAM_OTHERCALURL])) {
			$otherEvents = $eventReader->getEvents($_POST[PARAM_OTHERCALURL]);
		} else {
			$otherEvents = null;
		}
	} catch (Exception $e) {
		$errorMessage = $e;
		$gotException = true;
	}
	
	try {
		$docGenerator->generateDoc($mainEvents, $otherEvents, $organisationInfo);
	} catch (CopyFileException $e) {
		$errorMessage = $e;
		$gotException = true;
	} catch (CreateTemporaryFileException $e) {
		$errorMessage = $e;
		$gotException = true;
	} catch (Exception $e) {
		$errorMessage = $e;
		$gotException = true;
	}
	
	if (!$gotException && file_exists('../results/' . FILENAME)) {
		$contentType = 'Content-type: application/vnd.openxmlformats-officedocument.wordprocessingml.document;';
		header('HTTP/1.0 201 Created');
		header("Expires: Mon, 1 Apr 1974 05:00:00 GMT");
		header("Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT");
		header("Cache-Control: no-cache, must-revalidate");
		header("Pragma: no-cache");
		header($contentType);
		header("Content-Disposition: attachment; filename=" . FILENAME);
		readfile('../results/' . FILENAME);
	} else {
		header('HTTP/1.0 500 Internal Server Error');
		echo ("HTTP Error 500 : " . $errorMessage);
	}
} else {
	header('HTTP/1.0 400 Bad Request');
	echo ("HTTP Error 400 : Bad Request");
}
