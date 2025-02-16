<?php
class TemplateConfig extends ibPlugin {
    public function __construct() {
		parent::__construct();
        // Create or open the SQLite database
        $this->createSecurityAssessmentTemplateTable();

		if (!is_dir($this->getDir()['Files'].'/templates')) {
			mkdir($this->getDir()['Files'].'/templates', 0755, true);
		}
    }

    private function createSecurityAssessmentTemplateTable() {
        // Create template table if it doesn't exist
        $this->sql->exec("CREATE TABLE IF NOT EXISTS security_assessment_templates (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          Status TEXT,
          FileName TEXT,
          TemplateName TEXT,
          Description TEXT,
          ThreatActorSlide INTEGER,
          Orientation TEXT,
          Created DATE,
          Updated DATE
        )");
    }

    public function getSecurityAssessmentTemplateConfigs() {
        $stmt = $this->sql->prepare("SELECT * FROM security_assessment_templates");
        $stmt->execute();
        $templates = $stmt->fetchAll(PDO::FETCH_ASSOC);
        return $templates;
    }

    public function getSecurityAssessmentTemplateConfigById($id) {
        $stmt = $this->sql->prepare("SELECT * FROM security_assessment_templates WHERE id = :id");
        $stmt->execute([':id' => $id]);

        $templates = $stmt->fetch(PDO::FETCH_ASSOC);
        if ($templates) {
          return $templates;
        } else {
          return false;
        }
    }

    public function getSecurityAssessmentActiveTemplate() {
        $stmt = $this->sql->prepare("SELECT * FROM security_assessment_templates WHERE Status = :Status");
        $stmt->execute([':Status' => 'Active']);

        $template = $stmt->fetchAll(PDO::FETCH_ASSOC);
        if ($template) {
          return $template;
        } else {
          return false;
        }
    }

    public function newSecurityAssessmentTemplateConfig($Status,$FileName,$TemplateName,$Description,$ThreatActorSlide,$Orientation) {
        try {
            // Check if filename already exists
            $checkStmt = $this->sql->prepare("SELECT COUNT(*) FROM security_assessment_templates WHERE FileName = :FileName OR TemplateName = :TemplateName");
            $checkStmt->execute([':FileName' => $FileName, ':TemplateName' => $TemplateName]);
            if ($checkStmt->fetchColumn() > 0) {
				$this->api->setAPIResponse('Error','Template Name already exists');
				return false;
            }
        } catch (PDOException $e) {
			$this->api->setAPIResponse('Error',$e);
        }
        $stmt = $this->sql->prepare("INSERT INTO security_assessment_templates (Status, FileName, TemplateName, Description, ThreatActorSlide, Orientation, Created) VALUES (:Status, :FileName, :TemplateName, :Description, :ThreatActorSlide, :Orientation, :Created)");
        try {
            $CurrentDate = new DateTime();
            $stmt->execute([':Status' => urldecode($Status), ':FileName' => urldecode($FileName), ':TemplateName' => urldecode($TemplateName), ':Description' => urldecode($Description), ':ThreatActorSlide' => urldecode($ThreatActorSlide), ':Orientation' => urldecode($Orientation), ':Created' => $CurrentDate->format('Y-m-d H:i:s')]);
            $id = $this->sql->lastInsertId();
            // Mark other templates as inactive
            // if ($Status == 'Active') {
            //     $statusStmt = $this->sql->prepare("SELECT id FROM security_assessment_templates WHERE Status == :Status");
            //     $statusStmt->execute([':Status' => 'Active']);
            //     $ActiveTemplates = $statusStmt->fetchAll(PDO::FETCH_ASSOC);
            //     foreach ($ActiveTemplates as $AT) {
            //         $setStatusStmt = $this->sql->prepare("UPDATE security_assessment_templates SET Status = :Status WHERE id == :id AND id != :thisid");
            //         $setStatusStmt->execute([':Status' => 'Inactive',':id' => $AT['id'],':thisid' => $id]);
            //     }
            // }
            $this->logging->writeLog("Templates","Created New Security Assessment Template","info");
			$this->api->setAPIResponseMessage('Template added successfully');
        } catch (PDOException $e) {
			$this->api->setAPIResponse('Error',$e);
        }
    }

    public function setSecurityAssessmentTemplateConfig($id,$Status,$FileName,$TemplateName,$Description,$ThreatActorSlide,$Orientation) {
        $templateConfig = $this->getSecurityAssessmentTemplateConfigById($id);
        if ($templateConfig) {
            if ($FileName !== null || $TemplateName !== null) {
                try {
                    // Check if new filename/template name already exists
                    $checkStmt = $this->sql->prepare("SELECT COUNT(*) FROM security_assessment_templates WHERE (FileName = :FileName OR TemplateName = :TemplateName) AND id != :id");
                    $checkStmt->execute([':FileName' => $FileName, ':TemplateName' => $TemplateName, ':id' => $id]);
                    if ($checkStmt->fetchColumn() > 0) {
						$this->api->setAPIResponse('Error','Template name already exists');
						return false;
                    }
                } catch (PDOException $e) {
					$this->api->setAPIResponse('Error',$e);
                }
            }

            $prepare = [];
            $execute = [];
            $execute[':id'] = $id;
            if ($Status !== null) {
                // if ($Status == 'Active') {
                //     $statusStmt = $this->sql->prepare("SELECT id FROM security_assessment_templates WHERE Status == :Status");
                //     $statusStmt->execute([':Status' => 'Active']);
                //     $ActiveTemplates = $statusStmt->fetchAll(PDO::FETCH_ASSOC);
                //     foreach ($ActiveTemplates as $AT) {
                //         $setStatusStmt = $this->sql->prepare("UPDATE security_assessment_templates SET Status = :Status WHERE id == :id AND id != :thisid");
                //         $setStatusStmt->execute([':Status' => 'Inactive',':id' => $AT['id'],':thisid' => $id]);
                //     }
                // }
                $prepare[] = 'Status = :Status';
                $execute[':Status'] = urldecode($Status);
            }
            if ($FileName !== null) {
                $prepare[] = 'FileName = :FileName';
                $execute[':FileName'] = urldecode($FileName);
            }
            if ($TemplateName !== null) {
                $prepare[] = 'TemplateName = :TemplateName';
                $execute[':TemplateName'] = urldecode($TemplateName);
            }
            if ($Description !== null) {
                $prepare[] = 'Description = :Description';
                $execute[':Description'] = urldecode($Description);
            }
            if ($ThreatActorSlide !== null) {
                $prepare[] = 'ThreatActorSlide = :ThreatActorSlide';
                $execute[':ThreatActorSlide'] = urldecode($ThreatActorSlide);
            }
            if ($Orientation !== null) {
                $prepare[] = 'Orientation = :Orientation';
                $execute[':Orientation'] = urldecode($Orientation);
            }
            $stmt = $this->sql->prepare('UPDATE security_assessment_templates SET '.implode(", ",$prepare).' WHERE id = :id');
            $stmt->execute($execute);
            if ($FileName !== null) {
                $uploadDir = $this->getDir()['Files'].'/templates/';
                if ($templateConfig['FileName']) {
                    if (file_exists($uploadDir.$templateConfig['FileName'])) {
                        if (!unlink($uploadDir.$templateConfig['FileName'])) {
							$this->api->setAPIResponse('Error','Failed to delete old template file');
							return false;
                        }
                    }
                }
            }
            $this->logging->writeLog("Templates","Updated Security Assessment Template: ".$TemplateName,"info");
			$this->api->setAPIResponseMessage('Template updated successfully');
        } else {
			$this->api->setAPIResponse('Error','Template does not exist');
        }
    }

    public function removeSecurityAssessmentTemplateConfig($id) {
        $templateConfig = $this->getSecurityAssessmentTemplateConfigById($id);
        if ($templateConfig) {
          $uploadDir = $this->getDir()['Files'].'/templates/';
          if ($templateConfig['FileName']) {
            if (file_exists($uploadDir.$templateConfig['FileName'])) {
                if (!unlink($uploadDir.$templateConfig['FileName'])) {
					$this->api->setAPIResponse('Error','Failed to delete template file');
					return false;
                }
            }
          }
          $stmt = $this->sql->prepare("DELETE FROM security_assessment_templates WHERE id = :id");
          $stmt->execute([':id' => $id]);
          if ($this->getSecurityAssessmentTemplateConfigById($id)) {
			$this->api->setAPIResponse('Error','Failed to delete template');
          } else {
            $this->logging->writeLog("Templates","Removed Security Assessment Template: ".$id,"warning");
			$this->api->setAPIResponseMessage('Template deleted successfully');
          }
        }
    }
}