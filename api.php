<?php

// Get Plugin Settings
$app->get('/plugin/ib/settings', function ($request, $response, $args) {
	$ibPlugin = new ibPlugin();
	if ($ibPlugin->auth->checkAccess('ADMIN-CONFIG')) {
		$ibPlugin->api->setAPIResponseData($ibPlugin->_pluginGetSettings());
	}
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// ** // ******************* // ** //
// ** // SECURITY ASSESSMENT // ** //
// ** // ******************* // ** //

// Generate Security Assessment
$app->post('/plugin/ib/assessment/security/generate', function ($request, $response, $args) {
	$ibPlugin = new SecurityAssessment();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-SECURITYASSESSMENT']) ?? null) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        if ((isset($data['APIKey']) OR isset($_COOKIE['crypt'])) AND isset($data['StartDateTime']) AND isset($data['EndDateTime']) AND isset($data['Realm']) AND isset($data['id']) AND isset($data['templates'])) {
            if ($ibPlugin->SetCSPConfiguration($data['APIKey'] ?? null,$data['Realm'] ?? null)) {
                if (isValidUuid($data['id'])) {

                    $config = [
                        'APIKey' => $data['APIKey'] ?? null,
                        'Realm' => $data['Realm'],
                        'StartDateTime' => $data['StartDateTime'],
                        'EndDateTime' => $data['EndDateTime'],
                        'UUID' => $data['id'],
                        'Templates' => $data['templates'],
                        'unnamed' => $data['unnamed'] ?? false,
                        'substring' => $data['substring'] ?? false,
                        'unknown' => $data['unknown'] ?? false,
                        'allTAInMetrics' => $data['allTAInMetrics'] ?? false
                    ];

                    $ibPlugin->generateSecurityReport($config);
                }
            }
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Get Assessment Progress
$app->get('/plugin/ib/assessment/security/progress', function ($request, $response, $args) {
	$ibPlugin = new SecurityAssessment();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-SECURITYASSESSMENT']) ?? null) {
        $data = $request->getQueryParams();
        if (isset($data['id']) AND isValidUuid($data['id'])) {
            $ibPlugin->api->setAPIResponseData($ibPlugin->getProgress($data['id'])); // Produces percentage for use on progress bar
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Download Security Assessment Report
$app->get('/plugin/ib/assessment/security/download', function ($request, $response, $args) {
    $ibPlugin = new SecurityAssessment();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-SECURITYASSESSMENT']) ?? null) {
        $data = $request->getQueryParams();
        if (isset($data['id']) AND isValidUuid($data['id'])) {
            $ibPlugin->logging->writeLog("Assessment","Downloaded security assessment report","info");
            $progressFile = $ibPlugin->getDir()['Files'].'/reports/report-'.$data['id'].'.progress';
            // Ensure the progress file exists and is readable
            if (file_exists($progressFile) && is_readable($progressFile)) {
                // Read the progress file content
                $progressData = json_decode(file_get_contents($progressFile), true);
                if (isset($progressData['Templates']) && is_array($progressData['Templates'])) {
                    $templates = $progressData['Templates'];
                    if (count($templates) === 1) {
                        // Single template file
                        $file = $ibPlugin->getDir()['Files'].'/reports/report-'.$data['id'].'-'.$templates[0];
                        if (file_exists($file) && is_readable($file)) {
                            $fileContent = file_get_contents($file);
                            $response->getBody()->write($fileContent);
                            return $response
                                ->withHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
                                ->withHeader('Content-Disposition', 'attachment; filename="reports-'.$data['id'].'-'.$templates[0].'"')
                                ->withHeader('Content-Transfer-Encoding', 'binary')
                                ->withHeader('Accept-Ranges', 'bytes')
                                ->withStatus($GLOBALS['responseCode']);
                        } else {
                            $ibPlugin->api->setAPIResponse('Error','Template file not found or not readable');
                        }
                    } elseif (count($templates) > 1) {
                        // Multiple template files, compress them into a zip file
                        $zip = new ZipArchive();
                        $zipFile = $ibPlugin->getDir()['Files'].'/reports/templates-'.$data['id'].'.zip';
                        
                        if ($zip->open($zipFile, ZipArchive::CREATE) === TRUE) {
                            foreach ($templates as $template) {
                                $file = $ibPlugin->getDir()['Files'].'/reports/report-'.$data['id'].'-'.$template;
                                if (file_exists($file) && is_readable($file)) {
                                    $zip->addFile($file, basename($file));
                                }
                            }
                            $zip->close();
                            $zipContent = file_get_contents($zipFile);
                            $response->getBody()->write($zipContent);
                            return $response
                                ->withHeader('Content-Type', 'application/zip')
                                ->withHeader('Content-Disposition', 'attachment; filename="reports-'.$data['id'].'.zip"')
                                ->withHeader('Content-Transfer-Encoding', 'binary')
                                ->withHeader('Accept-Ranges', 'bytes')
                                ->withStatus($GLOBALS['responseCode']);
                        } else {
                            $ibPlugin->api->setAPIResponse('Error','Failed to create zip file');
                        }
                    } else {
                        $ibPlugin->api->setAPIResponse('Error','No templates found');
                    }
                } else {
                    $ibPlugin->api->setAPIResponse('Error','Invalid progress data');
                }
            } else {
                $ibPlugin->api->setAPIResponse('Error','Progress file not found or not readable');
            }
        } else {
            $ibPlugin->api->setAPIResponse('Error','Invalid ID');
        }
    }
    $response->getBody()->write(jsonE($GLOBALS['api']));
    return $response
        ->withHeader('Content-Type', 'application/json;charset=UTF-8')
        ->withStatus($GLOBALS['responseCode']);
});

// ** // SECURITY ASSESSMENT TEMPLATES // ** //

// Get Security Assessment Templates
$app->get('/plugin/ib/assessment/security/config', function ($request, $response, $args) {
	$ibPlugin = new TemplateConfig();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        $ibPlugin->api->setAPIResponseData($ibPlugin->getSecurityAssessmentTemplateConfigs());
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// New Security Assessment Template
$app->post('/plugin/ib/assessment/security/config', function ($request, $response, $args) {
	$ibPlugin = new TemplateConfig();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        if (isset($data['TemplateName'])) {
            $Status = $data['Status'] ?? null;
            $FileName = $data['FileName'] ? $data['FileName'] . '.pptx' : null;
            $Description = $data['Description'] ?? null;
            $ThreatActorSlide = $data['ThreatActorSlide'] ?? null;
            $Orientation = $data['Orientation'] ?? null;
            $isDefault = $data['isDefault'] ?? null;
            $TemplateName = $data['TemplateName'];
            $ibPlugin->newSecurityAssessmentTemplateConfig($Status,$FileName,$TemplateName,$Description,$ThreatActorSlide,$Orientation,$isDefault);
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Update Security Assessment Template
$app->patch('/plugin/ib/assessment/security/config/{id}', function ($request, $response, $args) {
	$ibPlugin = new TemplateConfig();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        $Status = $data['Status'] ?? null;
        $FileName = $data['FileName'] ? $data['FileName'] . '.pptx' : null;
        $TemplateName = $data['TemplateName'] ?? null;
        $Orientation = $data['Orientation'] ?? null;
        $Description = $data['Description'] ?? null;
        $isDefault = $data['isDefault'] ?? null;
        $ThreatActorSlide = $data['ThreatActorSlide'] ?? null;
        $ibPlugin->setSecurityAssessmentTemplateConfig($args['id'],$Status,$FileName,$TemplateName,$Description,$ThreatActorSlide,$Orientation,$isDefault);
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Delete Security Assessment Template
$app->delete('/plugin/ib/assessment/security/config/{id}', function ($request, $response, $args) {
	$ibPlugin = new TemplateConfig();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $ibPlugin->removeSecurityAssessmentTemplateConfig($args['id']);
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Upload New Security Assessment Template
$app->post('/plugin/ib/assessment/security/config/upload', function ($request, $response, $args) {
    $ibPlugin = new ThreatActors();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $uploadedFiles = $request->getUploadedFiles();
        $postData = $request->getParsedBody();
        $uploadDir = $ibPlugin->getDir()['Files'].'/templates/'; // Define your upload directory

        // Handle PPTX image upload
        if (isset($uploadedFiles['pptx']) && $uploadedFiles['pptx']->getError() == UPLOAD_ERR_OK) {
            if (isset($postData['TemplateName'])) {
                $pptxFileName = basename($uploadedFiles['pptx']->getClientFilename());
                $pptxFilePath = $uploadDir . 'security-' . urldecode($postData['TemplateName']) . '.pptx';

                if (isValidFileType($pptxFileName, ['pptx'])) {
                    // Move the uploaded file to the designated directory
                    $uploadedFiles['pptx']->moveTo($pptxFilePath);
                    $ibPlugin->api->setAPIResponseMessage("Successfully uploaded PPTX file: $pptxFileName");
                } else {
                    $ibPlugin->api->setAPIResponse("Errors","Invalid PPTX File: $pptxFileName");
                }
            } else {
                $ibPlugin->api->setAPIResponse("Errors","PPTX File Name Missing");
            }
        }
    }

	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// ** // **************** // ** //
// ** // CLOUD ASSESSMENT // ** //
// ** // **************** // ** //

// Generate Cloud Assessment
$app->post('/plugin/ib/assessment/cloud/generate', function ($request, $response, $args) {
	$ibPlugin = new CloudAssessment();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CLOUDASSESSMENT']) ?? null) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        if ((isset($data['APIKey']) OR isset($_COOKIE['crypt'])) AND isset($data['Realm']) AND isset($data['id']) AND isset($data['templates'])) {
            if ($ibPlugin->SetCSPConfiguration($data['APIKey'] ?? null,$data['Realm'] ?? null)) {
                if (isValidUuid($data['id'])) {

                    $config = [
                        'APIKey' => $data['APIKey'] ?? null,
                        'Realm' => $data['Realm'],
                        'UUID' => $data['id'],
                        'Templates' => $data['templates']
                    ];

                    $ibPlugin->generateCloudReport($config);
                }
            }
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Get Assessment Progress
$app->get('/plugin/ib/assessment/cloud/progress', function ($request, $response, $args) {
	$ibPlugin = new CloudAssessment();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CLOUDASSESSMENT']) ?? null) {
        $data = $request->getQueryParams();
        if (isset($data['id']) AND isValidUuid($data['id'])) {
            $ibPlugin->api->setAPIResponseData($ibPlugin->getProgress($data['id'])); // Produces percentage for use on progress bar
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Download Security Assessment Report
$app->get('/plugin/ib/assessment/cloud/download', function ($request, $response, $args) {
    $ibPlugin = new CloudAssessment();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CLOUDASSESSMENT']) ?? null) {
        $data = $request->getQueryParams();
        if (isset($data['id']) AND isValidUuid($data['id'])) {
            $ibPlugin->logging->writeLog("Assessment","Downloaded cloud assessment report","info");
            $progressFile = $ibPlugin->getDir()['Files'].'/reports/report-'.$data['id'].'.progress';
            // Ensure the progress file exists and is readable
            if (file_exists($progressFile) && is_readable($progressFile)) {
                // Read the progress file content
                $progressData = json_decode(file_get_contents($progressFile), true);
                if (isset($progressData['Templates']) && is_array($progressData['Templates'])) {
                    $templates = $progressData['Templates'];
                    if (count($templates) === 1) {
                        // Single template file
                        $file = $ibPlugin->getDir()['Files'].'/reports/report-'.$data['id'].'-'.$templates[0];
                        if (file_exists($file) && is_readable($file)) {
                            $fileContent = file_get_contents($file);
                            $response->getBody()->write($fileContent);
                            return $response
                                ->withHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
                                ->withHeader('Content-Disposition', 'attachment; filename="reports-'.$data['id'].'-'.$templates[0].'"')
                                ->withHeader('Content-Transfer-Encoding', 'binary')
                                ->withHeader('Accept-Ranges', 'bytes')
                                ->withStatus($GLOBALS['responseCode']);
                        } else {
                            $ibPlugin->api->setAPIResponse('Error','Template file not found or not readable');
                        }
                    } elseif (count($templates) > 1) {
                        // Multiple template files, compress them into a zip file
                        $zip = new ZipArchive();
                        $zipFile = $ibPlugin->getDir()['Files'].'/reports/templates-'.$data['id'].'.zip';
                        
                        if ($zip->open($zipFile, ZipArchive::CREATE) === TRUE) {
                            foreach ($templates as $template) {
                                $file = $ibPlugin->getDir()['Files'].'/reports/report-'.$data['id'].'-'.$template;
                                if (file_exists($file) && is_readable($file)) {
                                    $zip->addFile($file, basename($file));
                                }
                            }
                            $zip->close();
                            $zipContent = file_get_contents($zipFile);
                            $response->getBody()->write($zipContent);
                            return $response
                                ->withHeader('Content-Type', 'application/zip')
                                ->withHeader('Content-Disposition', 'attachment; filename="reports-'.$data['id'].'.zip"')
                                ->withHeader('Content-Transfer-Encoding', 'binary')
                                ->withHeader('Accept-Ranges', 'bytes')
                                ->withStatus($GLOBALS['responseCode']);
                        } else {
                            $ibPlugin->api->setAPIResponse('Error','Failed to create zip file');
                        }
                    } else {
                        $ibPlugin->api->setAPIResponse('Error','No templates found');
                    }
                } else {
                    $ibPlugin->api->setAPIResponse('Error','Invalid progress data');
                }
            } else {
                $ibPlugin->api->setAPIResponse('Error','Progress file not found or not readable');
            }
        } else {
            $ibPlugin->api->setAPIResponse('Error','Invalid ID');
        }
    }
    $response->getBody()->write(jsonE($GLOBALS['api']));
    return $response
        ->withHeader('Content-Type', 'application/json;charset=UTF-8')
        ->withStatus($GLOBALS['responseCode']);
});

// ** // CLOUD ASSESSMENT TEMPLATES // ** //

// Get Cloud Assessment Templates
$app->get('/plugin/ib/assessment/cloud/config', function ($request, $response, $args) {
	$ibPlugin = new TemplateConfig();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        $ibPlugin->api->setAPIResponseData($ibPlugin->getCloudAssessmentTemplateConfigs());
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// New Cloud Assessment Template
$app->post('/plugin/ib/assessment/cloud/config', function ($request, $response, $args) {
	$ibPlugin = new TemplateConfig();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        if (isset($data['TemplateName'])) {
            $Status = $data['Status'] ?? null;
            $FileName = $data['FileName'] ? $data['FileName'] . '.pptx' : null;
            $Description = $data['Description'] ?? null;
            $Orientation = $data['Orientation'] ?? null;
            $isDefault = $data['isDefault'] ?? null;
            $TemplateName = $data['TemplateName'];
            $ibPlugin->newCloudAssessmentTemplateConfig($Status,$FileName,$TemplateName,$Description,$Orientation,$isDefault);
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Update Cloud Assessment Template
$app->patch('/plugin/ib/assessment/cloud/config/{id}', function ($request, $response, $args) {
	$ibPlugin = new TemplateConfig();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        $Status = $data['Status'] ?? null;
        $FileName = $data['FileName'] ? $data['FileName'] . '.pptx' : null;
        $TemplateName = $data['TemplateName'] ?? null;
        $Orientation = $data['Orientation'] ?? null;
        $Description = $data['Description'] ?? null;
        $isDefault = $data['isDefault'] ?? null;
        $ibPlugin->setCloudAssessmentTemplateConfig($args['id'],$Status,$FileName,$TemplateName,$Description,$Orientation,$isDefault);
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Delete Cloud Assessment Template
$app->delete('/plugin/ib/assessment/cloud/config/{id}', function ($request, $response, $args) {
	$ibPlugin = new TemplateConfig();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $ibPlugin->removeCloudAssessmentTemplateConfig($args['id']);
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Upload New Cloud Assessment Template
$app->post('/plugin/ib/assessment/cloud/config/upload', function ($request, $response, $args) {
    $ibPlugin = new ThreatActors();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $uploadedFiles = $request->getUploadedFiles();
        $postData = $request->getParsedBody();
        $uploadDir = $ibPlugin->getDir()['Files'].'/templates/'; // Define your upload directory

        // Handle PPTX image upload
        if (isset($uploadedFiles['pptx']) && $uploadedFiles['pptx']->getError() == UPLOAD_ERR_OK) {
            if (isset($postData['TemplateName'])) {
                $pptxFileName = basename($uploadedFiles['pptx']->getClientFilename());
                $pptxFilePath = $uploadDir . 'cloud-' . urldecode($postData['TemplateName']) . '.pptx';

                if (isValidFileType($pptxFileName, ['pptx'])) {
                    // Move the uploaded file to the designated directory
                    $uploadedFiles['pptx']->moveTo($pptxFilePath);
                    $ibPlugin->api->setAPIResponseMessage("Successfully uploaded PPTX file: $pptxFileName");
                } else {
                    $ibPlugin->api->setAPIResponse("Errors","Invalid PPTX File: $pptxFileName");
                }
            } else {
                $ibPlugin->api->setAPIResponse("Errors","PPTX File Name Missing");
            }
        }
    }

	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Get Assessment Tracking Records
$app->get('/plugin/ib/assessment/reports/records', function ($request, $response, $args) {
	$ibPlugin = new AssessmentReporting();
    $data = $request->getQueryParams();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-REPORTING']) ?: 'ACL-REPORTING') {
        if (isset($data['granularity']) && isset($data['filters'])) {
            $Filters = $data['filters'];
            $Start = $data['start'] ?? null;
            $End = $data['end'] ?? null;
            $ibPlugin->logging->writeLog("Reporting","Queried Assessment Reports","info");
            $ibPlugin->api->setAPIResponseData($ibPlugin->getAssessmentReports($data['granularity'],json_decode($Filters,true),$Start,$End));
        } else {
            $ibPlugin->api->setAPIResponse('Error','Required values are missing from the request');
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Get Assessment Tracking Stats
$app->get('/plugin/ib/assessment/reports/stats', function ($request, $response, $args) {
	$ibPlugin = new AssessmentReporting();
    $data = $request->getQueryParams();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-REPORTING']) ?: 'ACL-REPORTING') {
        if (isset($data['granularity']) && isset($data['filters'])) {
            $Filters = $data['filters'];
            $Start = $data['start'] ?? null;
            $End = $data['end'] ?? null;
            $ibPlugin->logging->writeLog("Reporting","Queried Assessment Reports","info");
            $ibPlugin->api->setAPIResponseData($ibPlugin->getAssessmentReportsStats($_REQUEST['granularity'],json_decode($Filters,true),$Start,$End));
        } else {
            $ibPlugin->api->setAPIResponse('Error','Required values are missing from the request');
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Get Assessment Tracking Summary
$app->get('/plugin/ib/assessment/reports/summary', function ($request, $response, $args) {
	$ibPlugin = new AssessmentReporting();
    $data = $request->getQueryParams();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-REPORTING']) ?: 'ACL-REPORTING') {
        $ibPlugin->logging->writeLog("Reporting","Queried Assessment Reports","info");
        $ibPlugin->api->setAPIResponseData($ibPlugin->getAssessmentReportsSummary());
    }

	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Get Threat Actor List (IB Portal)
$app->post('/plugin/ib/threatactors', function ($request, $response, $args) {
	$ibPlugin = new ThreatActors();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-THREATACTORS']) ?? null) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        if ($ibPlugin->SetCSPConfiguration($data['APIKey'] ?? null,$data['Realm'] ?? null)) {
            $ibPlugin->getThreatActors($data);
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Get Threat Actor By ID (IB Portal)
$app->post('/plugin/ib/threatactor/{ActorID}', function ($request, $response, $args) {
	$ibPlugin = new ThreatActors();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-THREATACTORS']) ?? null) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        if ($ibPlugin->SetCSPConfiguration($data['APIKey'] ?? null,$data['Realm'] ?? null)) {
            $ibPlugin->GetB1ThreatActor($args['ActorID'],$data['Page'] ?? null);
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Get Configured Threat Actors
$app->get('/plugin/ib/threatactors/config', function ($request, $response, $args) {
	$ibPlugin = new ThreatActors();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        $ibPlugin->api->setAPIResponseData($ibPlugin->getThreatActorConfigs());
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// New Configured Threat Actor
$app->post('/plugin/ib/threatactors/config', function ($request, $response, $args) {
	$ibPlugin = new ThreatActors();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        if (isset($data['name'])) {
            $SVG = $data['SVG'] ? $data['SVG'] . '.svg' : null;
            $PNG = $data['PNG'] ? $data['PNG'] . '.png' : null;
            $URLStub = $data['URLStub'] ?? null;
            $ibPlugin->newThreatActorConfig($data['name'],$SVG,$PNG,$URLStub);
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Update Configured Threat Actor
$app->patch('/plugin/ib/threatactors/config/{id}', function ($request, $response, $args) {
	$ibPlugin = new ThreatActors();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        $Name = $data['name'] ?? null;
        $SVG = $data['SVG'] ? $data['SVG'] . '.svg' : null;
        $PNG = $data['PNG'] ? $data['PNG'] . '.png' : null;
        $URLStub = $data['URLStub'] ?? null;
        $ibPlugin->setThreatActorConfig($args['id'],$Name,$SVG,$PNG,$URLStub);
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Delete Configured Threat Actor
$app->delete('/plugin/ib/threatactors/config/{id}', function ($request, $response, $args) {
	$ibPlugin = new ThreatActors();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $ibPlugin->removeThreatActorConfig($args['id']);
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Upload Image(s) For Configured Threat Actor
$app->post('/plugin/ib/threatactors/config/upload', function ($request, $response, $args) {
    $ibPlugin = new ThreatActors();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG')) {
        $uploadedFiles = $request->getUploadedFiles();
        $postData = $request->getParsedBody();
        $uploadDir = $ibPlugin->getDir()['Assets'].'/images/Threat Actors/Uploads/'; // Define your upload directory

        $ResponseData = [
            'Items' => [],
            'Errors' => []
        ];

        // Handle SVG image upload
        if (isset($uploadedFiles['svgImage']) && $uploadedFiles['svgImage']->getError() == UPLOAD_ERR_OK) {
            if (isset($postData['svgFileName'])) {
                $svgFileName = basename($uploadedFiles['svgImage']->getClientFilename());
                $svgFilePath = $uploadDir . urldecode($postData['svgFileName']) . '.svg';

                if (isValidFileType($svgFileName, ['svg'])) {
                    // Move the uploaded file to the designated directory
                    $uploadedFiles['svgImage']->moveTo($svgFilePath);
                    $ResponseData['Items'][] = "SVG image uploaded successfully: $svgFileName";
                } else {
                    $ResponseData['Errors'][] = "Invalid SVG File: $svgFileName";
                }
            } else {
                $ResponseData['Errors'][] = "SVG File Name Missing";
            }
        }

        // Handle PNG image upload
        if (isset($uploadedFiles['pngImage']) && $uploadedFiles['pngImage']->getError() == UPLOAD_ERR_OK) {
            if (isset($postData['pngFileName'])) {
                $pngFileName = basename($uploadedFiles['pngImage']->getClientFilename());
                $pngFilePath = $uploadDir . urldecode($postData['pngFileName']) . '.png';

                if (isValidFileType($pngFileName, ['png'])) {
                    // Move the uploaded file to the designated directory
                    $uploadedFiles['pngImage']->moveTo($pngFilePath);
                    $ResponseData['Items'][] = "PNG image uploaded successfully: $pngFileName";
                } else {
                    $ResponseData['Errors'][] = "Invalid PNG File: $pngFileName";
                }
            } else {
                $ResponseData['Errors'][] = "PNG File Name Missing";
            }
        }
        $ibPlugin->api->setAPIResponseData($ResponseData);
    }

	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Generate License Assessment
$app->post('/plugin/ib/assessment/license/generate', function ($request, $response, $args) {
	$ibPlugin = new LicenseAssessment();
    if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-LICENSEUSAGE']) ?? null) {
        $data = $ibPlugin->api->getAPIRequestData($request);
        if ($ibPlugin->SetCSPConfiguration($data['APIKey'] ?? null,$data['Realm'] ?? null)) {
            if ((isset($data['APIKey']) OR isset($_COOKIE['crypt'])) AND isset($data['StartDateTime']) AND isset($data['EndDateTime']) AND isset($data['Realm'])) {
                $ibPlugin->getLicenseCount($data['StartDateTime'],$data['EndDateTime'],$data['Realm']);
            }
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

// Lame Delegation
$app->get('/plugin/ib/sittingducks/lame', function ($request, $response, $args) {
	$ibPlugin = new SittingDucks("one.one.one.one");
    if ($ibPlugin->auth->checkAccess('ADMIN-CONFIG') ?? null) {
        $data = $request->getQueryParams();
        if (isset($data['domain'])) {
            $ibPlugin->api->setAPIResponseData([$ibPlugin->isLame($data['domain'])]);
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});
// Sitting Ducks
$app->get('/plugin/ib/sittingducks/check', function ($request, $response, $args) {
	$ibPlugin = new SittingDucks("one.one.one.one");
    if ($ibPlugin->auth->checkAccess('ADMIN-CONFIG') ?? null) {
        $data = $request->getQueryParams();
        if (isset($data['domain'])) {

            // Example usage
            $result = array(
                'Dangling' => $ibPlugin->checkDNSRecords($data['domain']),
                'Expired' => $ibPlugin->checkDomainExpiry($data['domain']),
                'Subdomains' => $ibPlugin->checkSubdomainTakeover($data['domain'])
            );
            
            $ibPlugin->api->setAPIResponseData($result);
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});
// Sitting Ducks - Subdomains
$app->get('/plugin/ib/sittingducks/subdomains', function ($request, $response, $args) {
	$ibPlugin = new SittingDucks("one.one.one.one");
    if ($ibPlugin->auth->checkAccess('ADMIN-CONFIG') ?? null) {
        $data = $request->getQueryParams();
        if (isset($data['domain'])) {
        
            # Generate permutations and alterations of subdomain names
            $subdomains = ["sub1.example.com", "sub2.example.com"]; // Example subdomains

            $result = array(
                'cnameRecords' => $cnameRecords,
                // 'permutations' => $ibPlugin->generatePermutations($subdomains),
                'links' => $ibPlugin->searchWebServer($data['domain']),
                'ctLogs' => $ibPlugin->getCertificateTransparencyLogs($data['domain']),
                'sslCert' => $ibPlugin->checkSSLCertificates($data['domain']),
                'dnsRecords' => $ibPlugin->getDNSRecords($data['domain'])
            );
            
            $ibPlugin->api->setAPIResponseData($ibPlugin->extractUniqueDomains($result));
        }
    }
	$response->getBody()->write(jsonE($GLOBALS['api']));
	return $response
		->withHeader('Content-Type', 'application/json;charset=UTF-8')
		->withStatus($GLOBALS['responseCode']);
});

