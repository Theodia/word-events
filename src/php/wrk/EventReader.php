<?php
/**
 * Created by PhpStorm.
 * User: morattelp
 * Date: 19.03.2019
 * Time: 15:40
 */

require_once '../../vendor/autoload.php';
include_once 'beans/Event.php';


/**
 * Class IcalReader Permet de récupérer des évènements d'un calendrier électronique dans des objets Event.
 */
class EventReader {
	
	private $icalParser;
	private $dateFormat;
	
	/**
	 * IcalReader constructor.
	 * @param String $dateFormat Format de la date choisi par l'utilisateur.
	 */
	public function __construct($dateFormat) {
		$this->icalParser = new om\IcalParser();
		$this->dateFormat = $dateFormat;
	}
	
	/**
	 * Utilise la classe IcalParser pour parser le fichier ical, crée un objet Event pour chaque évènement du calendrier
	 * et renvoie le tableau contenant les objets Event.
	 * @param String $icalURL URL menant à un calendrier éléctronique contenant les évènements à lire.
	 * @return array Tableau contenant la liste des évènements 'timeEvents' et la liste des évènements du jour 'dateEvents'.
	 * @throws Exception
	 */
	public function getEvents($icalURL) {
		
		$this->icalParser->parseFile($icalURL);
		$eventsData = $this->icalParser->getSortedEvents();
		$dateEvents = [];
		$timeEvents = [];
		foreach ($eventsData as $eventData) {
			$locationTab = explode(', ', $eventData['LOCATION']);
			$location = $locationTab[0];
			
			$newEvent = new Event($this->dateFormat, $eventData['DTSTART'], $location, $eventData['SUMMARY'], $eventData['DESCRIPTION']);

			if ($eventData['isDateEvent']){
				$dateEvents[] = $newEvent;
			} else {
				$timeEvents[] = $newEvent;
			}
		}
		return [
			'timeEvents' => $timeEvents,
			'dateEvents' => $dateEvents
		];
	}
	
}