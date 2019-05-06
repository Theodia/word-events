<?php
/**
 * Created by PhpStorm.
 * User: morattelp
 * Date: 19.03.2019
 * Time: 15:43
 */

/**
 * Class Event
 * Classe d'objet Event, initialisant et regroupant les informations concernant un événement.
 */
class Event {

	
	public $date;
    public $time;
    public $location;
    public $summary;
    public $description;
    public $descriptionInOXML;
	
	/**
	 * Event constructor.
	 * @param String $locale Code de la région linguistique.
	 * @param String $datePattern Paterne de formatage de la date.
	 * @param String $timePattern Paterne de formatage de l'heure.
	 * @param DateTime $dateTime Date et heure de l'évènement non-formatée.
	 * @param String $location Lieu de l'événement.
	 * @param String $summary Résumé de l'événement.
	 * @param String $description Description de l'événement.
	 */
	public function __construct($locale, $datePattern, $timePattern, $dateTime, $location = '', $summary = '', $description = '') {
		//Initialize attributes
		$this->location = $location;
		$this->summary = $summary;
		$this->description = $description;
		$this->initDescriptionInOXML();
		
		//create date and time formatters
		$dateFormatter = $this->createDateTimeFormat($datePattern, $locale);
		$timeFormatter = $this->createDateTimeFormat($timePattern, $locale);
		
		//format $dateTime into $date and $time
		$this->date = $dateFormatter->format($dateTime);
		$this->time = $timeFormatter->format($dateTime);
	}
	
	/**
	 *Initialise l'attribut 'descriptionInHTML' en convertissant la description de l'HTML à l'OXML.
	 */
	private function initDescriptionInOXML(){
		$htmlParser = new HTMLtoOpenXML\Parser();

		$value = $htmlParser->fromHTML($this->description);
		$pos1 = strpos($value, '<w:p>');
		$value = substr_replace($value, '', $pos1, 5);
		$pos2 = strpos($value, '</w:p>');
		$this->descriptionInOXML = substr_replace($value, '', $pos2, 6);
		
	}
	
	/**
	 * Crée un objet IntlDateFormatter avec le paterne donné et l'attribut locale de l'Event
	 * @param String $pattern Paterne pour définir le formatage d'objet DateTime
	 * @param String $locale Code de la région linguistique.
	 * @return IntlDateFormatter Formateur d'objet DateTime
	 */
	private function createDateTimeFormat($pattern, $locale){
		$dateTimeFormatter = new IntlDateFormatter(
			$locale,
			IntlDateFormatter::NONE,
			IntlDateFormatter::NONE,
			null, null, $pattern
		);
		return $dateTimeFormatter;
	}
}