<?php
/**
 * Created by PhpStorm.
 * User: morattelp
 * Date: 20.03.2019
 * Time: 07:59
 */

use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
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
     * @throws CopyFileException
     * @throws CreateTemporaryFileException
     * @throws \PhpOffice\PhpWord\Exception\Exception
     * @throws Exception
     */
	public function generateDoc($mainEvents, $otherEvents, $organisationInfo) {
		$infosKeys = array_keys($organisationInfo);
		$templateProcessor = new PhpOffice\PhpWord\TemplateProcessor($this->templateURL);
		$initialDocVariables = $templateProcessor->getVariables();
        if ($otherEvents !== null){
        	
            $events = $otherEvents['timeEvents'];
	        
        } else {
            $events = $mainEvents['timeEvents'];
        }

        $dateEvents = $mainEvents['dateEvents'];
        
        $isDateEventVariable = in_array('t.event.eventoftheday', $initialDocVariables);
		
		if (in_array('t.table.nocellmerge', $initialDocVariables)) {
            $this->processTableNoCellMerge($templateProcessor, $events);

		} elseif (in_array('t.table.cellmerge.date', $initialDocVariables)) {
            $this->processTableCellMerge($templateProcessor, $isDateEventVariable, $dateEvents, $events);
		}
		
		foreach ($initialDocVariables as $docVariable) {
			
			if (in_array($docVariable, $infosKeys)) {
				
				if ($docVariable === 't.organisation.logo'){
					$templateProcessor->setImageValue($docVariable, $organisationInfo[$docVariable]);
				} else {
					$templateProcessor->setValue($docVariable, $organisationInfo[$docVariable]);
				}
			}
		}
		
        foreach ($templateProcessor->getVariables() as $variable) {
            if (substr($variable, 0, 8) === 't.table.' || substr($variable, 0, 21) === 't.event.eventoftheday'){
                $templateProcessor->setValue($variable, '');
            }
        }
        $templateProcessor->saveAs('../results/fiche_dominicale.docx');
	}
	
	/**
	 * Génère le tableau sans cellule fusionnée
	 * @param TemplateProcessor $templateProcessor
	 * @param $events
	 * @throws Exception
	 */
	private function processTableNoCellMerge($templateProcessor, $events){
		$templateProcessor->cloneRow('t.table.nocellmerge', sizeof($events));
		
		for($i = 0; $i < sizeof($events); $i++) {
			$event = $events[$i];
			$eventNum = $i + 1;
			$variables = [
				't.event.date#' . $eventNum,
				't.event.location#' . $eventNum,
				't.event.time#' . $eventNum,
				't.event.summary#' . $eventNum,
				't.event.description#' . $eventNum
			];
			
			$this->setEventValues($event, $variables, $templateProcessor);
			
		}
	}
	
	/**
	 * Génère le tableau avec des cellules fusionnées
	 * @param TemplateProcessor $templateProcessor
	 * @param $isDateEventVariable
	 * @param $dateEvents
	 * @param $events
	 * @throws Exception
	 */
	private function processTableCellMerge($templateProcessor, $isDateEventVariable, $dateEvents, $events){
		$days = $this->groupByDays($events);
		$templateProcessor->cloneRow('t.table.cellmerge.date', sizeof($days));
		
		for ($i = 0; $i < sizeof($days); $i++){
			$date = $days[$i][0]->date;
			$dateNum = $i + 1;
			$dateVariable = 't.event.date#' . $dateNum;
			
			if ($isDateEventVariable){
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
				$templateProcessor->setValue('t.event.eventoftheday#' . $dateNum, $totalSummary);
			}
			
			$templateProcessor->setValue($dateVariable, $date);
			$templateProcessor->cloneRow('t.table.cellmerge.event#' . $dateNum, sizeof($days[$i]));
			
			for($j = 0; $j < sizeof($days[$i]); $j++){
				$event = $days[$i][$j];
				$eventNum = $j + 1;
				$variables = [
					't.event.location#' . $dateNum . '#' . $eventNum,
					't.event.time#' . $dateNum . '#' . $eventNum,
					't.event.summary#' . $dateNum . '#' . $eventNum,
					't.event.description#' . $dateNum . '#' . $eventNum
				];
				
				$this->setEventValues($event, $variables, $templateProcessor);
			}
			
		}
	}

    /**
     * Groupe les événements par date
     * @param Event[] $events
     * @return Event[][]
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
     * remplace les balises d'événements par les valeurs des événements donnés
     * @param Event $event
     * @param string[] $variables
     * @param PhpOffice\PhpWord\TemplateProcessor $templateProcessor
     * @throws Exception
     */
    private function setEventValues($event, $variables, $templateProcessor){
        foreach ($variables as $variable) {
            $valueTypeTab = explode('#', $variable);
            $valueType = $valueTypeTab[0];

            switch ($valueType){
                case 't.event.date':
                    $value = $event->date;
                    break;
                case 't.event.time':
                    $value = $event->time;
                    break;
                case 't.event.location':
                    $value = $event->location;
                    break;
                case 't.event.summary':
                    $value = $event->summary;
                    break;
                case 't.event.description':
	                $value = $event->descriptionInOXML;
                    break;
                default :
                    $value = null;
            }
            if ($value !== null){
            	if (in_array($variable, $templateProcessor->getVariables())){
            		if (strpos($variable, 't.event.description') !== false){
			            \PhpOffice\PhpWord\Settings::setOutputEscapingEnabled(false);
		            }
		            $templateProcessor->setValue($variable, $value);
		            \PhpOffice\PhpWord\Settings::setOutputEscapingEnabled(true);
	            }
            } else {
                throw new Exception("The variable '" . $variable . "', value '". $value ."' given didn't correspond to any Event attribute");
            }
        }
    }
}