<?php

use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;

include_once 'tags.php';

const TIME_EVENTS = "timeEvents";
const DATE_EVENTS = "dateEvents";

const PARAM_TEMPLATE_URL = "templateURL";
const PARAM_MAIN_CAL_URL = "calendarMainURL";
const PARAM_OTHER_CAL_URL = "calendarOtherURL";
const PARAM_LOGO_URL = "logoURL";

const PARAM_DATE_BEGIN = "dateBegin";
const PARAM_DATE_END = "dateEnd";

const PARAM_ORG_NAME = "organisationName";
const PARAM_ORG_LOCATION = "organisationLocation";
const PARAM_ORG_ADDRESS = "organisationAddress";
const PARAM_ORG_PHONE = "organisationPhoneNbr";

const PARAM_LOCALE = "locales";
const PARAM_DATE_PATTERN = "datePattern";
const PARAM_TIME_PATTERN = "timePattern";

const FILE_NAME = "fiche_dominicale.docx";
const RESULT_DIR = "../results/";

require_once '../../vendor/autoload.php';

include_once 'wrk/EventReader.php';
include_once 'wrk/DocGenerator.php';

//test si les paramètres POST obligatoires sont non-vides
if (!empty($_POST[PARAM_TEMPLATE_URL]) && !empty($_POST[PARAM_MAIN_CAL_URL]) && !empty($_POST[PARAM_DATE_PATTERN]) && !empty($_POST[PARAM_TIME_PATTERN])) {
	
	$errorMessage = "HTTP Error 500 : Internal Server Error";
	$gotException = false;
	if (file_exists(RESULT_DIR . FILE_NAME)){
		unlink(RESULT_DIR . FILE_NAME);
	}
	
	//Création des classes EventReader et DocGenerator
	$eventReader = new EventReader();
	$docGenerator = new DocGenerator($_POST[PARAM_TEMPLATE_URL]);
	
	//Récupération des informations de l'organisation données en paramètres
	$organisationInfo = [
		TAG_ORG_LOGO => $_POST[PARAM_LOGO_URL],
		TAG_ORG_NAME => $_POST[PARAM_ORG_NAME],
		TAG_ORG_LOCATION => $_POST[PARAM_ORG_LOCATION],
		TAG_ORG_ADDRESS => $_POST[PARAM_ORG_ADDRESS],
		TAG_ORG_PHONE => $_POST[PARAM_ORG_PHONE],
	];
	
	$icalTab = [$_POST[PARAM_MAIN_CAL_URL]];
	if (!empty($_POST[PARAM_OTHER_CAL_URL])) {
		$icalTab[] = $_POST[PARAM_OTHER_CAL_URL];
	}
	
	//Appel de la classe EventReader pour lire les événements et créer les objets Event
	try {
		$mainEvents = $eventReader->getEvents($icalTab, $_POST[PARAM_LOCALE], $_POST[PARAM_DATE_PATTERN], $_POST[PARAM_TIME_PATTERN], $_POST[PARAM_DATE_BEGIN], $_POST[PARAM_DATE_END]);
		
	} catch (Exception $e) {
		$errorMessage = $e->getMessage();
		$errorLine = $e->getLine();
		$errorFile = $e->getFile();
		$gotException = true;
	}
	
	if (!$gotException){
		//Appel de la classe DocGenerator pour générer le document avec les données des Event et les paramètres donnés
		try {
			$docGenerator->generateDoc($mainEvents, $organisationInfo, RESULT_DIR . FILE_NAME);
		} catch (CopyFileException $e) {
			$errorMessage = $e->getMessage();
			$errorLine = $e->getLine();
			$errorFile = $e->getFile();
			$gotException = true;
		} catch (CreateTemporaryFileException $e) {
			$errorMessage = $e->getMessage();
			$errorLine = $e->getLine();
			$errorFile = $e->getFile();
			$gotException = true;
		} catch (Exception $e) {
			$errorMessage = $e->getMessage();
			$errorLine = $e->getLine();
			$errorFile = $e->getFile();
			$gotException = true;
		}
	}
	
	//Si on a pas reçu d'Exception et que le document de résultat existe, on crée et envoie l'en-tête et on lit le fichier
	if (!$gotException && file_exists(RESULT_DIR . FILE_NAME)) {
		$contentType = 'Content-type: application/vnd.openxmlformats-officedocument.wordprocessingml.document;';
		header('HTTP/1.0 201 Created');
		header("Expires: Mon, 1 Apr 1974 05:00:00 GMT");
		header("Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT");
		header("Cache-Control: no-cache, must-revalidate");
		header("Pragma: no-cache");
		header($contentType);
		header("Content-Disposition: attachment; filename=" . FILE_NAME);
		readfile(RESULT_DIR . FILE_NAME);
	} else {
		header('HTTP/1.0 500 Internal Server Error');
		echo ("HTTP Error 500 : " . $errorMessage . " dans le fichier " . substr($errorFile, 47)  . " à la ligne " . $errorLine);
	}
} else {
	header('HTTP/1.0 400 Bad Request');
	echo ("HTTP Error 400 : Bad Request");
}
