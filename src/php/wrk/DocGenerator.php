<?php
/**
 * Created by PhpStorm.
 * User: morattelp
 * Date: 20.03.2019
 * Time: 07:59
 */

use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\TemplateProcessor;

require_once '../../vendor/autoload.php';

/**
 * Class DocGenerator Permet de générer un document Word sur la base d'un template .docx donné.
 */
class DocGenerator {
	
	private $phpWord;
	private $templateURL;
	
	/**
	 * DocGenerator constructor.
	 * @param string $templateURL URL du template ".docx" qui sera utilisé pour générer la fiche dominicale finale.
	 */
	public function __construct($templateURL) {
		$this->templateURL = $templateURL;
		$this->phpWord = new PhpOffice\PhpWord\PhpWord();

	}
	
	/**
	 * Prend des tableaux d'événements et crée un document .docx avec les données des événements et les données de la
	 * Paroisse.
	 * @param Event[][] $mainEvents Tableau contenant les objets Event des événements lithuriques (messes, ...).
	 * @param Event[][] $otherEvents Tableau contenant les objets Event des autres événements (concerts, ...).
	 * @param string[] $organisationInfo Tableau contenant les informations en String de la Paroisse.
	 * @param String $fileName Chemin et nom du fichier de résultat à sauver.
	 * @throws CopyFileException Si une erreur survient durant la copie du fichier template.
	 * @throws CreateTemporaryFileException Si une erreur survient durant la création du fichier temporaire.
	 * @throws \PhpOffice\PhpWord\Exception\Exception Si une erreur survient durant le remplacement d'un tag par une image ou dans la sauvegarde du fichier.
	 * @throws Exception Si une erreur survient dans le clonage des cellules de tableau ou dans le remplacement des tags.
	 */
	public function generateDoc($mainEvents, $otherEvents, $organisationInfo, $fileName) {
		$infosKeys = array_keys($organisationInfo);
		$templateProcessor = new PhpOffice\PhpWord\TemplateProcessor($this->templateURL);
		$initialDocVariables = $templateProcessor->getVariables();
        if ($otherEvents !== null){
        	
            $events = $otherEvents[TIME_EVENTS];
	        
        } else {
            $events = $mainEvents[TIME_EVENTS];
        }

        $dateEvents = $mainEvents[DATE_EVENTS];
        
		if (in_array(TAG_TABLE_NO_CELLMERGE, $initialDocVariables)) {
            $this->processTableNoCellMerge($templateProcessor, $events);

		} elseif (in_array(TAG_TABLE_CELLMERGE_DATE, $initialDocVariables)) {
            $this->processTableCellMerge($templateProcessor, $dateEvents, $events);
		}
		
		foreach ($initialDocVariables as $docVariable) {
			
			if (in_array($docVariable, $infosKeys)) {
				
				if ($docVariable === TAG_ORG_LOGO){
					$templateProcessor->setImageValue($docVariable, $organisationInfo[$docVariable]);
				} else {
					$templateProcessor->setValue($docVariable, $organisationInfo[$docVariable]);
				}
			}
		}
		
        foreach ($templateProcessor->getVariables() as $variable) {
            if (substr($variable, 0, 8) === substr(TAG_TABLE_NO_CELLMERGE, 0, 8) || substr($variable, 0, 21) === TAG_EVENT_EVENTOFTHEDAY){
                $templateProcessor->setValue($variable, '');
            }
        }
		$templateProcessor->saveAs($fileName);
	}
	
	/**
	 * Génère le tableau sans cellule fusionnée.
	 * @param TemplateProcessor $templateProcessor Instance d'une classe de phpWord servant à modifier les fichiers docx.
	 * @param Event[] $events Tableau d'objets Event, contenant les information des événements.
	 * @throws Exception Si une erreur survient dans les méthodes cloneRow() ou setEventValues().
	 */
	private function processTableNoCellMerge($templateProcessor, $events){
		$templateProcessor->cloneRow(TAG_TABLE_NO_CELLMERGE, sizeof($events));
		
		for($i = 0; $i < sizeof($events); $i++) {
			$event = $events[$i];
			$eventNum = $i + 1;
			$variables = [
				TAG_EVENT_DATE . TAG_SEPARATOR . $eventNum,
				TAG_EVENT_LOCATION . TAG_SEPARATOR . $eventNum,
				TAG_EVENT_TIME . TAG_SEPARATOR . $eventNum,
				TAG_EVENT_SUMMARY . TAG_SEPARATOR . $eventNum,
				TAG_EVENT_DESCRIPTION . TAG_SEPARATOR . $eventNum
			];
			
			$this->setEventValues($event, $variables, $templateProcessor);
			
		}
	}
	
	/**
	 * Génère le tableau avec des cellules fusionnées
	 * @param TemplateProcessor $templateProcessor Instance d'une classe de phpWord servant à modifier les fichiers docx.
	 * @param Event[] $dateEvents tableau d'objets Event des événements durant toute une journée (événements spéciaux).
	 * @param Event[] $events tableau d'objets Event des événements normaux.
	 * @throws Exception Si une erreur survient dans les méthodes cloneRow() ou setEventValues().
	 */
	private function processTableCellMerge($templateProcessor, $dateEvents, $events){
		$days = $this->groupByDays($events);
		$templateProcessor->cloneRow(TAG_TABLE_CELLMERGE_DATE, sizeof($days));
		
		for ($i = 0; $i < sizeof($days); $i++){
			$date = $days[$i][0]->date;
			$dateNum = $i + 1;
			$dateVariable = TAG_EVENT_DATE . TAG_SEPARATOR . $dateNum;
			
			if (in_array(TAG_EVENT_EVENTOFTHEDAY . TAG_SEPARATOR . $dateNum, $templateProcessor->getVariables())){
				$dateEventCount = 0;
				$totalSummary = '';
				
				foreach ($dateEvents as $dateEvent) {
					if ($date === $dateEvent->date){
						if ($dateEventCount !== 0){
							$totalSummary .= ', ';
						}
						$dateEventCount++;
						$totalSummary .= $dateEvent->summary;
					}
				}
				$templateProcessor->setValue(TAG_EVENT_EVENTOFTHEDAY . TAG_SEPARATOR . $dateNum, $totalSummary);
			}
			
			$templateProcessor->setValue($dateVariable, $date);
			$templateProcessor->cloneRow(TAG_TABLE_CELLMERGE_EVENT . TAG_SEPARATOR . $dateNum, sizeof($days[$i]));
			
			for($j = 0; $j < sizeof($days[$i]); $j++){
				$event = $days[$i][$j];
				$eventNum = $j + 1;
				$variables = [
					TAG_EVENT_LOCATION . TAG_SEPARATOR . $dateNum . TAG_SEPARATOR . $eventNum,
					TAG_EVENT_TIME . TAG_SEPARATOR . $dateNum . TAG_SEPARATOR . $eventNum,
					TAG_EVENT_SUMMARY . TAG_SEPARATOR . $dateNum . TAG_SEPARATOR . $eventNum,
					TAG_EVENT_DESCRIPTION . TAG_SEPARATOR . $dateNum . TAG_SEPARATOR . $eventNum
				];
				
				$this->setEventValues($event, $variables, $templateProcessor);
			}
			
		}
	}

    /**
     * Groupe les événements par date.
     * @param Event[] $events table des événements à ordrer.
     * @return Event[][] tableau des événements ordrés par date.
     */
    private function groupByDays($events){
	    $days = [];
	    $oldDate = null;
	    $i = 0;
        foreach ($events as $event) {

            if ($oldDate == $event->date){
                $days[$i][] = $event;
            } else {

                if ($oldDate !== null){
                    $i++;
                }
                $oldDate = $event->date;
                $day = [];
                $day[]= $event;
                $days[$i] = $day;
            }
	    }
        return $days;
    }

    /**
     * Remplace les balises d'événements par les valeurs des événements donnés.
     * @param Event $event objet d'événement qui contient les attributs d'un événement.
     * @param string[] $variables tableau des noms de tags.
     * @param PhpOffice\PhpWord\TemplateProcessor $templateProcessor Instance d'une classe de phpWord servant à modifier les fichiers docx.
     * @throws Exception si une des valeurs de la table variables ne correspond à aucun tag dans le switch.
     */
    private function setEventValues($event, $variables, $templateProcessor){
        foreach ($variables as $variable) {
            $valueTypeTab = explode(TAG_SEPARATOR, $variable);
            $valueType = $valueTypeTab[0];

            switch ($valueType){
                case TAG_EVENT_DATE:
                    $value = $event->date;
                    break;
                case TAG_EVENT_TIME:
                    $value = $event->time;
                    break;
                case TAG_EVENT_LOCATION:
                    $value = $event->location;
                    break;
                case TAG_EVENT_SUMMARY:
                    $value = $event->summary;
                    break;
                case TAG_EVENT_DESCRIPTION:
	                $value = $event->descriptionInOXML;
                    break;
                default :
                    $value = null;
            }
            if ($value !== null){
            	if (in_array($variable, $templateProcessor->getVariables())){
            		if (strpos($variable, TAG_EVENT_DESCRIPTION) !== false){
			            Settings::setOutputEscapingEnabled(false);
		            }
		            $templateProcessor->setValue($variable, $value);
		            Settings::setOutputEscapingEnabled(true);
	            }
            } else {
                throw new Exception("la variable '". $variable ."' ne correspond à aucun tag.");
            }
        }
    }
}