<?php
/**
 * Created by PhpStorm.
 * User: morattelp
 * Date: 19.03.2019
 * Time: 15:43
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
     * @param DateTime $dateTime
     * @param string $dateFormat
     * @param string $location
     * @param string $summary
     * @param string $description
     * @throws Exception
     */
	public function __construct($dateFormat, $dateTime, $location = '', $summary = '', $description = '') {

		$this->date = $dateTime->format($dateFormat);
		$this->time = $dateTime->format('H:i');
		$this->location = $location;
		$this->summary = $summary;
		$this->description = $description;
		$this->initDescriptionInOXML();
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
}