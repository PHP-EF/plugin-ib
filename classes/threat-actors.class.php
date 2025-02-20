<?php
class ThreatActors extends ibPortal {
	public function __construct() {
		parent::__construct();
		$this->createThreatActorTable();

		$imagesDir = $this->getDir()['Assets'].'/images/Threat Actors/';
		if (!is_dir($imagesDir)) {
			mkdir($imagesDir, 0755, true);
		}
		$uploadDir = $this->getDir()['Assets'].'/images/Threat Actors/Uploads/';
		if (!is_dir($uploadDir)) {
			mkdir($uploadDir, 0755, true);
		}
	}

    private function createThreatActorTable() {
        // Create users table if it doesn't exist
        $this->sql->exec("CREATE TABLE IF NOT EXISTS threat_actors (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          Name TEXT UNIQUE,
          SVG TEXT,
          PNG TEXT,
          URLStub TEXT
        )");
    }

    public function getThreatActorConfigById($id) {
        $stmt = $this->sql->prepare("SELECT * FROM threat_actors WHERE id = :id");
        $stmt->execute([':id' => $id]);
        $threatActors = $stmt->fetch(PDO::FETCH_ASSOC);
        if ($threatActors) {
          return $threatActors;
        } else {
          return false;
        }
    }

    public function getThreatActorConfigByName($name) {
        $stmt = $this->sql->prepare("SELECT * FROM threat_actors WHERE LOWER(Name) = LOWER(:name)");
        $stmt->execute([':name' => $name]);
        $threatActors = $stmt->fetch(PDO::FETCH_ASSOC);
        if ($threatActors) {
          return $threatActors;
        } else {
          return false;
        }
    }

    public function getThreatActorConfigs() {
        $stmt = $this->sql->prepare("SELECT * FROM threat_actors");
        $stmt->execute();
        $threatActors = $stmt->fetchAll(PDO::FETCH_ASSOC);
        return $threatActors;
    }

    public function newThreatActorConfig($Name,$SVG,$PNG,$URLStub) {
        if ($Name != "") {
            $ThreatActorConfig = $this->getThreatActorConfigs();
            try {
                // Check if filename already exists
                $checkStmt = $this->sql->prepare("SELECT COUNT(*) FROM threat_actors WHERE Name = :Name");
                $checkStmt->execute([':Name' => urldecode($Name)]);
                if ($checkStmt->fetchColumn() > 0) {
					$this->api->setAPIResponse('Error','Threat Actor Already Exists');
					return false;
                }
            } catch (PDOException $e) {
				$this->api->setAPIResponse('Error',$e);
            }
            $stmt = $this->sql->prepare("INSERT INTO threat_actors (Name, SVG, PNG, URLStub) VALUES (:Name, :SVG, :PNG, :URLStub)");
            try {
                $stmt->execute([':Name' => urldecode($Name), ':SVG' => urldecode($SVG), ':PNG' => urldecode($PNG), ':URLStub' => urldecode($URLStub)]);
                $this->logging->writeLog("ThreatActors","Created new Threat Actor: ".$Name,"info");
				$this->api->setAPIResponseMessage('Threat Actor added successfully');
            } catch (PDOException $e) {
				$this->api->setAPIResponse('Error',$e);
            }
        } else {
			$this->api->setAPIResponse('Error','Threat Actor name missing');
        }
    }

    public function setThreatActorConfig($id,$Name,$SVG,$PNG,$URLStub) {
        if ($Name != "") {
            $ThreatActorConfig = $this->getThreatActorConfigById($id);
            if ($ThreatActorConfig) {
                $prepare = [];
                $execute = [];
                $execute[':id'] = $id;
                if ($Name !== null) {
                    try {
                        // Check if new filename/template name already exists
                        $checkStmt = $this->sql->prepare("SELECT COUNT(*) FROM threat_actors WHERE Name = :Name AND id != :id");
                        $checkStmt->execute([':Name' => $Name, ':id' => $id]);
                        if ($checkStmt->fetchColumn() > 0) {
							$this->api->setAPIResponse('Error','Threat Actor name already exists');
							return false;
                        }
                    } catch (PDOException $e) {
						$this->api->setAPIResponse('Error',$e);
                    }
                    $prepare[] = 'Name = :Name';
                    $execute[':Name'] = urldecode($Name);
                }
                if ($SVG !== null) {
                    $prepare[] = 'SVG = :SVG';
                    $execute[':SVG'] = urldecode($SVG);
                }
                if ($PNG !== null) {
                    $prepare[] = 'PNG = :PNG';
                    $execute[':PNG'] = urldecode($PNG);
                }
                if ($URLStub !== null) {
                    $prepare[] = 'URLStub = :URLStub';
                    $execute[':URLStub'] = urldecode($URLStub);
                }
                $stmt = $this->sql->prepare('UPDATE threat_actors SET '.implode(", ",$prepare).' WHERE id = :id');
                $stmt->execute($execute);
                $this->logging->writeLog("ThreatActors","Updated Threat Actor: ".$Name,"info");
				$this->api->setAPIResponseMessage('Threat Actor updated successfully');
            } else {
				$this->api->setAPIResponse('Error','Threat Actor does not exist');
            }
        } else {
			$this->api->setAPIResponse('Error','Threat Actor name missing');
        }
    }

    public function removeThreatActorConfig($id) {
        $ThreatActorConfig = $this->getThreatActorConfigById($id);
        if ($ThreatActorConfig) {
          $uploadDir = $this->getDir()['Assets'].'/images/Threat Actors/Uploads/';
          if ($ThreatActorConfig['PNG']) {
            if (file_exists($uploadDir.$ThreatActorConfig['PNG'])) {
                if (!unlink($uploadDir.$ThreatActorConfig['PNG'])) {
					$this->api->setAPIResponse('Error','Failed to delete PNG file');
                    return false;
                }
            }
          }
          if ($ThreatActorConfig['SVG']) {
            if (file_exists($uploadDir.$ThreatActorConfig['SVG'])) {
                if (!unlink($uploadDir.$ThreatActorConfig['SVG'])) {
					$this->api->setAPIResponse('Error','Failed to delete SVG file');
                    return false;
                }
            }
          }
          $stmt = $this->sql->prepare("DELETE FROM threat_actors WHERE id = :id");
          $stmt->execute([':id' => $id]);
          if ($this->getThreatActorConfigById($id)) {
			$this->api->setAPIResponse('Error','Failed to delete Threat Actor');
			return false;
          } else {
            $this->logging->writeLog("ThreatActors","Removed Threat Actor: ".$id,"warning");
			$this->api->setAPIResponseMessage('Successfully deleted Threat Actor');
			return false;
          }
        }
    }

	// Function called from API
	public function GetThreatActors($data) {
		if ((isset($data['APIKey']) OR isset($_COOKIE['crypt'])) AND isset($data['StartDateTime']) AND isset($data['EndDateTime']) AND isset($data['Realm'])) {
			$UserInfo = $this->GetCSPCurrentUser();
			if (isset($UserInfo->result->name)) {
				$this->logging->writeLog("ThreatActors",$UserInfo->result->name." queried list of Threat Actors","info");
				$Actors = $this->GetB1ThreatActorIds($data['StartDateTime'],$data['EndDateTime']);
				if (!isset($Actors->Error)) {
					$this->api->setAPIResponseData($this->GetB1ThreatActorsById($Actors,$data['unnamed'],$data['substring']));
				} else {
					$this->api->setAPIResponse('Error','Unable to get list of Threat Actors','502',$Actors);
				};
			}
		} else {
			$this->api->setAPIResponse('Error','API Key, Start & End Date and Realm are required fields.');
		}
	}

	// Get list of Threat Actor IDs from CubeJS
	public function GetB1ThreatActorIds($StartDateTime,$EndDateTime) {
		$StartDimension = str_replace('Z','',$StartDateTime);
		$EndDimension = str_replace('Z','',$EndDateTime);
		// Workaround
		//$Actors = QueryCubeJS('{"segments":[],"timeDimensions":[{"dimension":"PortunusAggIPSummary.timestamp","granularity":null,"dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"ungrouped":false,"order":{"PortunusAggIPSummary.timestampMax":"desc"},"measures":["PortunusAggIPSummary.count"],"dimensions":["PortunusAggIPSummary.threat_indicator","PortunusAggIPSummary.actor_id"],"limit":1000,"filters":[{"and":[{"operator":"set","member":"PortunusAggIPSummary.threat_indicator"},{"operator":"set","member":"PortunusAggIPSummary.actor_id"}]}]}');
		$Actors = $this->QueryCubeJS('{"measures":[],"segments":[],"dimensions":["ThreatActors.storageid","ThreatActors.ikbactorid","ThreatActors.domain","ThreatActors.ikbfirstsubmittedts","ThreatActors.vtfirstdetectedts","ThreatActors.firstdetectedts","ThreatActors.lastdetectedts"],"timeDimensions":[{"dimension":"ThreatActors.lastdetectedts","granularity":null,"dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"ungrouped":false}');
		if (isset($Actors->result->data)) {
		  return $Actors->result->data;
		} else {
		  return $Actors;
		}
	  }
	
	public function GetB1ThreatActor($ActorID,$Page = 1) {
		$Results = $this->QueryCSP('get','/tide/threat-enrichment/clusterfox/actors/search?actor_id='.$ActorID.'&page='.$Page);
		if ($Results) {
			$this->api->setAPIResponseData($Results);
		} else {
			$this->api->setAPIResponse('Error','Unable to retrieve list of Threat Actors');
		}
	}
	  
	function GetB1ThreatActorsById($Actors,$unnamed,$substring,$unknown) {
		$ActorArr = json_decode(json_encode($Actors),true);
		$UniqueIds = array_unique(array_column($ActorArr, 'ThreatActors.ikbactorid'));
		$Results = array();
		$ActorInfo = array();
		$ArrayChunk = array_chunk($UniqueIds, 10);
		$Requests = [];
		foreach ($ArrayChunk as $Chunk) {
		  	$CsvString = implode(',',$Chunk);
		  	$Requests[] = $this->QueryCSPMultiRequestBuilder('get','/tide/threat-enrichment/clusterfox/actors/search?actor_id='.$CsvString);
		}
		$Responses = $this->QueryCSPMulti($Requests);
		foreach ($Responses as $Response) {
		  foreach ($Response['Body']->actors as $Actor) {
				$ObservedIOCKeys = array_keys(array_column($ActorArr, 'ThreatActors.ikbactorid'),$Actor->actor_id);
				$ObservedIOCCount = count($ObservedIOCKeys);
				$ObservedIOCs = [];
			foreach ($ObservedIOCKeys as $ObservedIOCKey) {
			  	$ObservedIOCs[] = $ActorArr[$ObservedIOCKey];
			}
			if ($Actor->actor_id != "" && $Actor->actor_name != "") {
				// Ignore Unnamed & Substring Actors
				$UnnamedActor = str_starts_with($Actor->actor_name,'unnamed_actor');
				$SubstringActor = str_starts_with($Actor->actor_name,'substring');
				$UnknownActor = str_starts_with($Actor->actor_name,'Unknown');
				if (($UnnamedActor && $unnamed == 'true') || ($SubstringActor && $substring == 'true') || ($UnknownActor && $unknown == 'true') || (!$UnnamedActor && !$SubstringActor && !$UnknownActor)) {
					$NewArr = array(
						'actor_id' => $Actor->actor_id,
						'actor_name' => $Actor->actor_name,
						'actor_description' => $Actor->actor_description,
						'related_count' => $Actor->related_count,
						'observed_count' => $ObservedIOCCount,
						'observed_iocs' => $ObservedIOCs
					);
					if (isset($Actor->external_references)) {
						$NewArr['external_references'] = $Actor->external_references;
					} else {
						$NewArr['external_references'] = [];
					}
					if (isset($Actor->infoblox_references)) {
						$NewArr['infoblox_references'] = $Actor->infoblox_references;
					} else {
						$NewArr['infoblox_references'] = [];
					}
					if (isset($Actor->purpose)) {
						$NewArr['purpose'] = $Actor->purpose;
					} else {
						$NewArr['purpose'] = [];
					}
					if (isset($Actor->ttp)) {
						$NewArr['ttp'] = $Actor->ttp;
					} else {
						$NewArr['ttp'] = [];
					}
						array_push($Results,$NewArr);
					}
				}
		  	}
		}
		return $Results;
	}

	function GetB1ThreatActorsByIdEU($Actors,$unnamed,$substring,$unknown) {
		$UniqueIds = array_unique(array_column($Actors, 'PortunusAggIPSummary.actor_id'));
		$Results = array();
		$ActorInfo = array();
		foreach ($UniqueIds as $UniqueId) {
		  // Workaround for problematic Threat Actors
		  // These timeout when using the 'batch_actor_summary_with_indicators' API Endpoint
		  $WorkaroundArr = array(
			//'c2303ad0-0f9e-4349-a71e-821794e202bd' // Revolver Rabbit - TIDE-850
		  );
		  if (in_array($UniqueId, $WorkaroundArr)) {
			$ActorQuery = QueryCSP('get','tide-ng-threat-actor/v1/actor?_filter=id=="'.$UniqueId.'" and page==1');
			if (isset($ActorQuery)) {
			  $NewArr = array(
				'actor_id' => $ActorQuery->actor_id,
				'actor_name' => $ActorQuery->actor_name,
				'actor_description' => $ActorQuery->actor_description,
				'related_count' => $ActorQuery->related_count,
				'related_indicators_with_dates' => null,
				'related_indicators' => $ActorQuery->related_indicators,
			  );
			  if (isset($ActorQuery->external_references)) {
				$NewArr['external_references'] = $ActorQuery->external_references;
			  } else {
				$NewArr['external_references'] = [];
			  }
			  if (isset($ActorQuery->infoblox_references)) {
				$NewArr['infoblox_references'] = $ActorQuery->infoblox_references;
			  } else {
				$NewArr['infoblox_references'] = [];
			  }
			  array_push($Results,$NewArr);
			}
		  // End of Workaround
		  } else {
			$Ids = array();
			$Ids[] = array_keys(array_column($Actors, 'PortunusAggIPSummary.actor_id'),$UniqueId);
			$Indicators = array();
			foreach ($Ids as $Id) {
			  foreach ($Id as $Idsub) {
				$Indicators[] = $Actors[$Idsub]->{'PortunusAggIPSummary.threat_indicator'};
			  }
			}
			$ActorInfo[] = array(
			  "actor_id" => $Actors[$Id[0]]->{'PortunusAggIPSummary.actor_id'},
			  "indicators" => $Indicators
			);
		  }
		}
	  
		$ArrayChunk = array_chunk($ActorInfo, 5);
		$Requests = [];
		foreach ($ArrayChunk as $Chunk) {
		  $Query = json_encode(array(
			'actor_indicators' => $Chunk
		  ));
		  $Requests[] = $this->QueryCSPMultiRequestBuilder('post','tide-ng-threat-actor/v1/batch_actor_summary_with_indicators',$Query);
		}
	  
		$Responses = $this->QueryCSPMulti($Requests);
		foreach ($Responses as $Response) {
		  if (isset($Response['Body']->actor_responses)) {
			foreach ($Response['Body']->actor_responses as $AR) {
			  // Ignore Unnamed & Substring Actors
			  $UnnamedActor = str_starts_with($AR->actor_name,'unnamed_actor');
			  $SubstringActor = str_starts_with($AR->actor_name,'substring');
			  $UnknownActor = str_starts_with($AR->actor_name,'unknown');
			  if (($UnnamedActor && $unnamed == 'true') || ($SubstringActor && $substring == 'true') || ($UnknownActor && $unknown == 'true') || (!$UnnamedActor && !$SubstringActor && !$UnknownActor)) {
				$NewArr = array(
					'actor_id' => $AR->actor_id,
					'actor_name' => $AR->actor_name,
					'actor_description' => $AR->actor_description,
					'related_count' => $AR->related_count,
					'related_indicators_with_dates' => $AR->related_indicators_with_dates,
					'related_indicators' => null,
				);
				if (isset($AR->external_references)) {
					$NewArr['external_references'] = $AR->external_references;
				} else {
					$NewArr['external_references'] = [];
				}
				if (isset($AR->infoblox_references)) {
					$NewArr['infoblox_references'] = $AR->infoblox_references;
				} else {
					$NewArr['infoblox_references'] = [];
				}
				array_push($Results,$NewArr);
			  }
			}
		  }
		}
		return $Results;
	}
}