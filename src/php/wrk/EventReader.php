<?php
/**
 * Created by PhpStorm.
 * User: morattelp
 * Date: 19.03.2019
 * Time: 15:40
 */

require_once '../../vendor/autoload.php';
include_once 'beans/Event.php';

const EVENT_START_DATE = 'DTSTART';
const EVENT_SUMMARY = 'SUMMARY';
const EVENT_LOCATION = 'LOCATION';
const EVENT_DESCRIPTION = 'DESCRIPTION';
const EVENT_IS_DATE_EVENT = 'isDateEvent';
/**
 * Class EventReader Permet de récupérer des évènements d'un calendrier électronique dans des objets Event.
 */
class EventReader {
	
	private $icalParser;
	
	/**
	 * EventReader constructor.
	 */
	public function __construct() {
		$this->icalParser = new om\IcalParser();
	}
	
	/**
	 * Utilise la classe IcalParser pour parser le fichier ical, crée un objet Event pour chaque évènement du calendrier
	 * et renvoie le tableau contenant les objets Event.
	 * @param String $icalURL URL menant à un calendrier éléctronique contenant les évènements à lire.
	 * @param String $locale Code de la région linguistique.
	 * @param String $datePattern Paterne de formatage de la date.
	 * @param String $timePattern Paterne de formatage de l'heure.
	 * @return array Tableau contenant la liste des évènements 'timeEvents' et la liste des évènements du jour 'dateEvents'.
	 * @throws Exception Si une erreur survient durant la lecture du fichier ical.
	 */
	public function getEvents($icalURL, $locale, $datePattern, $timePattern) {
		
		$this->icalParser->parseFile($icalURL);
		$eventsData = $this->icalParser->getSortedEvents();
		$dateEvents = [];
		$timeEvents = [];
		foreach ($eventsData as $eventData) {
			$locationTab = explode(', ', $eventData[EVENT_LOCATION]);
			$location = $locationTab[0];
			
			$newEvent = new Event($locale, $datePattern, $timePattern,  $eventData[EVENT_START_DATE], $location, $eventData[EVENT_SUMMARY], $eventData[EVENT_DESCRIPTION]);

			if ($eventData[EVENT_IS_DATE_EVENT]){
				$dateEvents[] = $newEvent;
			} else {
				$timeEvents[] = $newEvent;
			}
		}
		return [
			TIME_EVENTS => $timeEvents,
			DATE_EVENTS => $dateEvents
		];
	}
	
}