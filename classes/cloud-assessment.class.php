<?php

use Label305\PptxExtractor\Basic\BasicExtractor;
use Label305\PptxExtractor\Basic\BasicInjector;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class CloudAssessment extends ibPortal {
	private $AssessmentReporting;
	private $TemplateConfig;

	public function __construct() {
		parent::__construct();
		if (!is_dir($this->getDir()['Files'].'/reports')) {
			mkdir($this->getDir()['Files'].'/reports', 0755, true);
		}
		$this->AssessmentReporting = new AssessmentReporting();
		$this->TemplateConfig = new TemplateConfig();
	}

	public function generateCloudReport($config) {
		// Check Active Template Exists
		$SelectedTemplates = [];
		foreach ($config['Templates'] as $TemplateId) {
			$SelectedTemplate = $this->TemplateConfig->getCloudAssessmentTemplateConfigById($TemplateId);
			if ($SelectedTemplate) {
				$SelectedTemplates[] = $SelectedTemplate;
			}
		}
		if (empty($SelectedTemplates)) {
			$this->api->setAPIResponse('Error','No active template selected');
			return false;
		}

		// Create Progress File
		$this->createProgress($config['UUID'],$SelectedTemplates);

		// Check API Key is valid & get User Info
		$UserInfo = $this->GetCSPCurrentUser();
		if (is_array($UserInfo) && isset($UserInfo['Error'])) {
			$Status = $UserInfo['Status'];
			$Error = $UserInfo['Error'];
		} else {
			header('Content-Type: application/json; charset=utf-8');
			// echo json_encode(array(
			// 	'result' => 'Success',
			// 	'message' => 'Started'
			// ));
			// fastcgi_finish_request();
	
			// Logging / Reporting
			$AccountInfo = $this->QueryCSP("get","v2/current_user/accounts");
			$CurrentAccount = $AccountInfo->results[array_search($UserInfo->result->account_id, array_column($AccountInfo->results, 'id'))];
			$this->logging->writeLog("Assessment",$UserInfo->result->name." requested a cloud assessment report for: ".$CurrentAccount->name,"info");
			$ReportRecordId = $this->AssessmentReporting->newReportEntry('Cloud Assessment',$UserInfo->result->name,$CurrentAccount->name,$config['Realm'],$config['UUID'],"Started");
	
			// Set Progress
			$Progress = 0;
	
			// Set Time Dimensions
			// $StartDimension = str_replace('Z','',$config['StartDateTime']);
			// $EndDimension = str_replace('Z','',$config['EndDateTime']);

			// Get date/time 30 days earlier
			$pastDate = new DateTime();
			$pastDate->modify('-30 days');
			$StartDimension = $pastDate->format('Y-m-d\TH:i:s.v');

			// Get current date/time
			$currentDate = new DateTime();
			$EndDimension = $currentDate->format('Y-m-d\TH:i:s.v');

			$Progress = $this->writeProgress($config['UUID'],$Progress,"Collecting Metrics");
			$CubeJSRequests = array(
				'TotalAssetsCount' => '{"segments": [],"timeDimensions": [{"dimension": "AssetDetails.doc_updated_at","dateRange": ["'.$StartDimension.'","'.$EndDimension.'"],"granularity": null}],"ungrouped": false,"dimensions": [],"measures": ["AssetDetails.count"]}',
				'AssetsByClassification' => '{"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails.doc_asset_insight_classification","assetinsighttypes.label","AssetDetails.doc_asset_insight_sub_classification","assetinsightsubclassifications.label"],"filters":[{"member":"AssetDetails.is_valid_sub_classification","operator":"equals","values":["true"]},{"or":[{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["zombie"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["zombie"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["zombie/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]}]},{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["compliance"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["compliance"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["compliance/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]}]},{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["ghost"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["ghost"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["ghost/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["false"]}]}]}],"timezone":"UTC","segments":[]}',
				'AssetsByClassificationCount' => '{"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":[],"filters":[{"member":"AssetDetails.is_valid_sub_classification","operator":"equals","values":["true"]},{"or":[{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["zombie"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["zombie"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["zombie/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]}]},{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["compliance"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["compliance"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["compliance/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]}]},{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["ghost"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["ghost"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["ghost/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["false"]}]}]}],"timezone":"UTC","segments":[]}',
				'AssetsWithMissingRecords' => '{"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}]}],"dimensions":["assetinsightindicators.label","AssetDetails.doc_asset_insight_indicator"],"filters":[{"member":"AssetDetails.doc_asset_insight_indicator","operator":"startsWith","values":["registration"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]}],"timezone":"UTC","segments":[]}',
				'AssetsWithMissingRecordsCount' => '{"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":[],"filters":[{"member":"AssetDetails.doc_asset_insight_indicator","operator":"startsWith","values":["registration"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]}],"timezone":"UTC","segments":[]}',
				'CloudSubnetUtilizationAbove50' => '{"measures":["NetworkInsightsSubnet.count"],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]},{"operator":"gte","member":"NetworkInsightsSubnet.utilization_percent","values":["50"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.provider"],"timeDimensions":[{"dimension":"NetworkInsightsSubnet.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}]}',
				'CloudSubnetUtilizationBelow50' => '{"measures":["NetworkInsightsSubnet.count"],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]},{"operator":"lt","member":"NetworkInsightsSubnet.utilization_percent","values":["50"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.provider"],"timeDimensions":[{"dimension":"NetworkInsightsSubnet.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}]}',
				'CloudSubnetsByProvider' => '{"dimensions":["NetworkInsightsSubnet.provider"],"ungrouped":false,"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"dimension":"NetworkInsightsSubnet.updated_at","granularity":null}],"measures":["NetworkInsightsSubnet.count"],"segments":[]}',
				'CloudIPsByProvider' => '{"measures":["NetworkInsightsSubnet.count"],"segments":[],"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"dimension":"NetworkInsightsSubnet.updated_at","granularity":null}],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.provider","values":["AWS","GCP","Azure"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.utilization_used","NetworkInsightsSubnet.utilization_total","NetworkInsightsSubnet.provider"]}',
				'HighRiskDNSRecords' => '{"measures":["NetworkInsightsDnsRecords.count"],"timeDimensions":[{"dimension":"NetworkInsightsDnsRecords.evaluation_time","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["NetworkInsightsDnsRecords.indicator_id"],"filters":[],"timezone":"UTC","segments":[]}'
			);

			$CubeJSResults = $this->QueryCubeJSMulti($CubeJSRequests);
	
			// Extract Powerpoint Template(s) as Zip
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Extracting template(s)");
			foreach ($SelectedTemplates as &$SelectedTemplate) {
				$ExtractedDir = $this->getDir()['Files'].'/reports/'.str_replace('.pptx','', 'report'.'-'.$config['UUID'].'-'.$SelectedTemplate['FileName']);
				extractZip($this->getDir()['Files'].'/templates/'.$SelectedTemplate['FileName'],$ExtractedDir);
				$SelectedTemplate['ExtractedDir'] = $ExtractedDir;
			}

			// Define the embedded sheets with their corresponding file numbers
			// This needs to match across all active templates at this moment
			$EmbeddedSheets = [
				// 'TopDetectedProperties' => 4,
				// 'ContentFiltration' => 5,
				// 'DNSActivity' => 6,
				// 'DNSFirewallActivity' => 7,
				// 'InsightDistribution' => 8,
				// 'Lookalikes' => 9
			];

			// Function to get the full path of the file based on the sheet name
			function getEmbeddedSheetFilePath($sheetName, $embeddedDirectory, $embeddedFiles, $EmbeddedSheets) {
				if (isset($EmbeddedSheets[$sheetName])) {
					$fileIndex = $EmbeddedSheets[$sheetName];
					if (isset($embeddedFiles[$fileIndex])) {
						return $embeddedDirectory . $embeddedFiles[$fileIndex];
					}
				}
			}

			// ********************** //
			// ** Reusable Metrics ** //
			// ********************** //

			// High Risk Assets
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building High Risk Assets");
			$AssetsByClassificationCount = $CubeJSResults['AssetsByClassificationCount']['Body']->result->data[0]->{'AssetDetails.count'} ?? 0;
			$AssetsWithMissingRecordsCount = $CubeJSResults['AssetsWithMissingRecordsCount']['Body']->result->data[0]->{'AssetDetails.count'} ?? 0;
			$HighRiskAssetsCount = $AssetsByClassificationCount + $AssetsWithMissingRecordsCount;

			// Total Assets
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Total Assets");
			$TotalAssetsCount = $CubeJSResults['TotalAssetsCount']['Body']->result->data[0]->{'AssetDetails.count'} ?? 0;

			// Assets By Classification
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Assets by Classification");
			$GhostAssetsCount = 0;
			$ZombieAssetsCount = 0;
			$NonCompliantAssetsCount = 0;
			$AssetsByClassification = $CubeJSResults['AssetsByClassification']['Body']->result->data ?? array();
			foreach ($AssetsByClassification as $value) {
				switch ($value->{'AssetDetails.doc_asset_insight_classification'}) {
					case 'ghost':
						$GhostAssetsCount += $value->{'AssetDetails.count'};
						break;
					case 'zombie':
						$ZombieAssetsCount += $value->{'AssetDetails.count'};
						switch ($value->{'assetinsightsubclassifications.label'}) {
							case 'Idle':
								$ZombieIdleAssetsCount += $value->{'AssetDetails.count'};
								break;
							case 'Orphan':
								$ZombieOrphanedAssetsCount += $value->{'AssetDetails.count'};
								break;
							case 'Resource Utilization':
								$ZombieAssetsResourceUtilizationCount += $value->{'AssetDetails.count'};
								break;
						}
						break;
					case 'compliance':
						$NonCompliantAssetsCount += $value->{'AssetDetails.count'};
						break;
				}
			}
			if ($ZombieAssetsCount != 0) {
				$ZombieIdleAssetsPerc = ($ZombieIdleAssetsCount / $ZombieAssetsCount) * 100;
				$ZombieOrphanedAssetsPerc = ($ZombieOrphanedAssetsCount / $ZombieAssetsCount) * 100;
				$ZombieAssetsResourceUtilizationPerc = ($ZombieAssetsResourceUtilizationCount / $ZombieAssetsCount) * 100;
			} else {
				$ZombieIdleAssetsPerc = 0;
				$ZombieOrphanedAssetsPerc = 0;
				$ZombieAssetsResourceUtilizationPerc = 0;
			}

			// Subnet Utilization
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Subnet Utilization");
			$CloudSubnetUtilizationAbove50 = $CubeJSResults['CloudSubnetUtilizationAbove50']['Body']->result->data ?? array();
			$CloudSubnetUtilizationAbove50ByProvider = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Total" => 0];
			$CloudSubnetUtilizationAbove50Total = 0;
			foreach ($CloudSubnetUtilizationAbove50 as $value) {
				$CloudSubnetUtilizationAbove50ByProvider[$value->{'NetworkInsightsSubnet.provider'}] = $value->{'NetworkInsightsSubnet.count'} ?? 0;
				$CloudSubnetUtilizationAbove50Total += $value->{'NetworkInsightsSubnet.count'} ?? 0;
			}
			$CloudSubnetUtilizationBelow50 = $CubeJSResults['CloudSubnetUtilizationBelow50']['Body']->result->data ?? array();
			$CloudSubnetUtilizationBelow50ByProvider = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Total" => 0];
			$CloudSubnetUtilizationBelow50Total = 0;
			foreach ($CloudSubnetUtilizationBelow50 as $value) {
				$CloudSubnetUtilizationBelow50ByProvider[$value->{'NetworkInsightsSubnet.provider'}] = $value->{'NetworkInsightsSubnet.count'} ?? 0;
				$CloudSubnetUtilizationBelow50Total += $value->{'NetworkInsightsSubnet.count'} ?? 0;
			}

			// Subnet by Provider
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Subnets by Provider");
			$CloudSubnetsByProvider = $CubeJSResults['CloudSubnetsByProvider']['Body']->result->data ?? array();
			$AWSSubnetsCount = 0;
			$AzureSubnetsCount = 0;
			$GCPSubnetsCount = 0;
			$TotalSubnetsCount = 0;
			$TotalCloudSubnetsCount = 0;
			foreach ($CloudSubnetsByProvider as $value) {
				switch ($value->{'NetworkInsightsSubnet.provider'}) {
					case 'AWS':
						$AWSSubnetsCount += $value->{'NetworkInsightsSubnet.count'} ?? 0;
						$TotalCloudSubnetsCount += $value->{'NetworkInsightsSubnet.count'} ?? 0;
						break;
					case 'Azure':
						$AzureSubnetsCount += $value->{'NetworkInsightsSubnet.count'} ?? 0;
						$TotalCloudSubnetsCount += $value->{'NetworkInsightsSubnet.count'} ?? 0;
						break;
					case 'GCP':
						$GCPSubnetsCount += $value->{'NetworkInsightsSubnet.count'} ?? 0;
						$TotalCloudSubnetsCount += $value->{'NetworkInsightsSubnet.count'} ?? 0;
						break;
				}
				$TotalSubnetsCount += $value->{'NetworkInsightsSubnet.count'} ?? 0;
			}
			if ($TotalCloudSubnetsCount != 0) {
				$AWSSubnetsPercentage = ($AWSSubnetsCount / $TotalCloudSubnetsCount) * 100;
				$AzureSubnetsPercentage = ($AzureSubnetsCount / $TotalCloudSubnetsCount) * 100;
				$GCPSubnetsPercentage = ($GCPSubnetsCount / $TotalCloudSubnetsCount) * 100;
			} else {
				$AWSSubnetsPercentage = 0;
				$AzureSubnetsPercentage = 0;
				$GCPSubnetsPercentage = 0;
			}

			// IPs by Provider
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building IPs by Provider");
			$CloudIPsByProvider = $CubeJSResults['CloudIPsByProvider']['Body']->result->data ?? array();
			$CloudIPsByProviderTotal = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Total" => 0];
			$CloudIPsByProviderTotalUsed = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Total" => 0];
			foreach ($CloudIPsByProvider as $CloudIPsByProviderEntry) {
				$Provider = $CloudIPsByProviderEntry->{'NetworkInsightsSubnet.provider'};
				$CloudIPsByProviderTotal[$Provider] += (int)$CloudIPsByProviderEntry->{'NetworkInsightsSubnet.utilization_total'};
				$CloudIPsByProviderTotal["Total"] += (int)$CloudIPsByProviderEntry->{'NetworkInsightsSubnet.utilization_total'};
				$CloudIPsByProviderTotalUsed[$Provider] += (int)$CloudIPsByProviderEntry->{'NetworkInsightsSubnet.utilization_used'};
				$CloudIPsByProviderTotalUsed["Total"] += (int)$CloudIPsByProviderEntry->{'NetworkInsightsSubnet.utilization_used'};
			}

			// High Risk DNS Records
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building High Risk DNS Records");
			$HighRiskDNSRecords = $CubeJSResults['HighRiskDNSRecords']['Body']->result->data ?? array();
			$AbandonedDNSCount = 0;
			$UntrustedDNSCount = 0;
			$DanglingDNSCount = 0;
			
			foreach ($HighRiskDNSRecords as $record) {
				switch ($record->{'NetworkInsightsDnsRecords.indicator_id'}) {
					case 'Abandoned':
						$AbandonedDNSCount += $record->{'NetworkInsightsDnsRecords.count'};
						break;
					case 'Untrusted':
						$UntrustedDNSCount += $record->{'NetworkInsightsDnsRecords.count'};
						break;
					case 'Dangling':
						$DanglingDNSCount += $record->{'NetworkInsightsDnsRecords.count'};
						break;
				}
			}
			


			// Loop for each selected template
			foreach ($SelectedTemplates as &$SelectedTemplate) {
				$embeddedDirectory = $SelectedTemplate['ExtractedDir'].'/ppt/embeddings/';
				$embeddedFiles = array_values(array_diff(scandir($embeddedDirectory), array('.', '..')));
				$this->logging->writeLog("Assessment","Embedded Files List","debug",['Template' => $SelectedTemplate, 'Embedded Files' => $embeddedFiles]);
	
				// Open PPTX Presentation _rels XML
				$xml_rels = null;
				$xml_rels = new DOMDocument('1.0', 'utf-8');
				$xml_rels->formatOutput = true;
				$xml_rels->preserveWhiteSpace = false;
				$xml_rels->load($SelectedTemplate['ExtractedDir'].'/ppt/_rels/presentation.xml.rels');

				// Open PPTX Presentation XML
				$xml_pres = new DOMDocument('1.0', 'utf-8');
				$xml_pres->formatOutput = true;
				$xml_pres->preserveWhiteSpace = false;
				$xml_pres->load($SelectedTemplate['ExtractedDir'].'/ppt/presentation.xml');

				// Save Core XML Files
				$xml_rels->save($SelectedTemplate['ExtractedDir'].'/ppt/_rels/presentation.xml.rels');
				$xml_pres->save($SelectedTemplate['ExtractedDir'].'/ppt/presentation.xml');

				// Rebuild Powerpoint Template Zip(s)
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Stitching Powerpoint Template(s)");
				compressZip($this->getDir()['Files'].'/reports/report'.'-'.$config['UUID'].'-'.$SelectedTemplate['FileName'],$SelectedTemplate['ExtractedDir']);

				// Cleanup Extracted Zip(s)
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Cleaning up");
				//UNCOMMENT ME rmdirRecursive($SelectedTemplate['ExtractedDir']);

				// Extract Powerpoint Template Strings
				// ** Using external library to save re-writing the string replacement functions manually. Will probably pull this in as native code at some point.
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Extract Powerpoint Strings");
				$extractor = new BasicExtractor();
				$mapping = $extractor->extractStringsAndCreateMappingFile(
					$this->getDir()['Files'].'/reports/report'.'-'.$config['UUID'].'-'.$SelectedTemplate['FileName'],
					$SelectedTemplate['ExtractedDir'].'-extracted.pptx'
				);

				$Progress = $this->writeProgress($config['UUID'],$Progress,"Injecting Powerpoint Strings");
				##// Slide 2 / 45 - Title Page & Contact Page
				// Get & Inject Customer Name, Contact Name & Email
				$mapping = replaceTag($mapping,'#CUSTOMERNAME',$CurrentAccount->name);
				$mapping = replaceTag($mapping,'#DATE',date("jS F Y"));
				$StartDate = new DateTime($StartDimension);
				$EndDate = new DateTime($EndDimension);
				$mapping = replaceTag($mapping,'#DATESOFCOLLECTION',$StartDate->format("jS F Y").' - '.$EndDate->format("jS F Y"));
				$mapping = replaceTag($mapping,'#NAME',$UserInfo->result->name);
				$mapping = replaceTag($mapping,'#EMAIL',$UserInfo->result->email);

				##// Slide 5 - Executive Summary
				$mapping = replaceTag($mapping,'#TAG01',number_abbr($HighRiskAssetsCount)); // High-Risk Assets

				$mapping = replaceTag($mapping,'#TAG03',number_abbr($CloudSubnetUtilizationAbove50Total)); // Overutilized Subnets (>=50%)
				$mapping = replaceTag($mapping,'#TAG04',number_abbr($CloudSubnetUtilizationBelow50Total)); // Underutilized Subnets (<50%)
				$mapping = replaceTag($mapping,'#TAG06',number_abbr($DanglingDNSCount)); // High-Risk DNS Records - Dangling
				$mapping = replaceTag($mapping,'#TAG07',number_abbr($AbandonedDNSCount)); // High-Risk DNS Records - Abandoned
				$mapping = replaceTag($mapping,'#TAG08',number_abbr($UntrustedDNSCount)); // High-Risk DNS Records - Untrusted
				
				##// Slide 7 - Assets
				$mapping = replaceTag($mapping,'#TAG09',number_abbr($TotalAssetsCount)); // Total Assets
				$mapping = replaceTag($mapping,'#TAG10',number_abbr($HighRiskAssetsCount)); // High-Risk Assets
				$mapping = replaceTag($mapping,'#TAG11',number_abbr($ZombieAssetsCount)); // Zombie Assets Count
				$mapping = replaceTag($mapping,'#TAG12',number_abbr($AssetsWithMissingRecordsCount)); // Assets Missing Record(s)
				$mapping = replaceTag($mapping,'#TAG13',number_abbr($NonCompliantAssetsCount)); // Non-Compliant Assets

				##// Slide 8 - Assets
				$mapping = replaceTag($mapping,'#TAG14',number_abbr($ZombieAssetsCount)); // Zombie Assets Count
				$mapping = replaceTag($mapping,'#TAG15',number_abbr($ZombieIdleAssetsCount)); // Idle Zombie Assets Count
				$mapping = replaceTag($mapping,'#TAG16',number_abbr($ZombieIdleAssetsPerc)); // Idle Zombie Assets Percentage
				$mapping = replaceTag($mapping,'#TAG17',number_abbr($ZombieOrphanedAssetsCount)); // Orphaned Assets Count
				$mapping = replaceTag($mapping,'#TAG18',number_abbr($ZombieOrphanedAssetsPerc)); // Orphaned Assets Percentage

				##// Slide 9 - Assets
				$mapping = replaceTag($mapping,'#TAG19',number_abbr($AssetsWithMissingRecordsCount)); // Assets Missing Record(s)
				$mapping = replaceTag($mapping,'#TAG20',number_abbr($NonCompliantAssetsCount)); // Non-Compliant Assets

				##// Slide 11 - IP/Subnet Allocation
				$mapping = replaceTag($mapping,'#TAG21',number_abbr($TotalCloudSubnetsCount)); // Total Subnets (Box)
				$mapping = replaceTag($mapping,'#TAG22',number_abbr($TotalCloudSubnetsCount)); // Total Subnets (Centre)
				$mapping = replaceTag($mapping,'#TAG23',number_abbr($AWSSubnetsCount)); // AWS Subnet Count
				$mapping = replaceTag($mapping,'#TAG24',number_abbr($AWSSubnetsPercentage)); // AWS Subnet Percentage
				$mapping = replaceTag($mapping,'#TAG25',number_abbr($AzureSubnetsCount)); // Azure Subnet Count
				$mapping = replaceTag($mapping,'#TAG26',number_abbr($AzureSubnetsPercentage)); // Azure Subnet Percentage
				$mapping = replaceTag($mapping,'#TAG27',number_abbr($GCPSubnetsCount)); // GCP Subnet Count
				$mapping = replaceTag($mapping,'#TAG28',number_abbr($GCPSubnetsPercentage)); // GCP Subnet Percentage
				$mapping = replaceTag($mapping,'#TAG29',number_abbr($CloudIPsByProviderTotalUsed['Total'])); // Total Allocated Cloud IPs
				$mapping = replaceTag($mapping,'#TAG30',number_abbr($CloudIPsByProviderTotalUsed['Azure'])); // Total Allocated Azure IPs
				$mapping = replaceTag($mapping,'#TAG31',number_abbr($CloudIPsByProviderTotalUsed['AWS'])); // Total Allocated AWS IPs
				$mapping = replaceTag($mapping,'#TAG32',number_abbr($CloudIPsByProviderTotalUsed['GCP'])); // Total Allocated GCP IPs

				##// Slide 12 - IP/Subnet Allocation
				$mapping = replaceTag($mapping,'#TAG33',number_abbr($AWSSubnetsCount)); // Total AWS Subnets
				$mapping = replaceTag($mapping,'#TAG34',number_abbr($CloudSubnetUtilizationAbove50ByProvider['AWS'])); // AWS Subnets above 50% Utilization
				$mapping = replaceTag($mapping,'#TAG35',number_abbr($CloudSubnetUtilizationBelow50ByProvider['AWS'])); // AWS Subnets below 50% Utilization
				$mapping = replaceTag($mapping,'#TAG36',number_abbr($AzureSubnetsCount)); // Total Azure Subnets
				$mapping = replaceTag($mapping,'#TAG37',number_abbr($CloudSubnetUtilizationAbove50ByProvider['Azure'])); // Azure Subnets above 50% Utilization
				$mapping = replaceTag($mapping,'#TAG38',number_abbr($CloudSubnetUtilizationBelow50ByProvider['Azure'])); // Azure Subnets below 50% Utilization
				$mapping = replaceTag($mapping,'#TAG39',number_abbr($GCPSubnetsCount)); // Total GCP Subnets
				$mapping = replaceTag($mapping,'#TAG40',number_abbr($CloudSubnetUtilizationAbove50ByProvider['GCP'])); // GCP Subnets above 50% Utilization
				$mapping = replaceTag($mapping,'#TAG41',number_abbr($CloudSubnetUtilizationBelow50ByProvider['GCP'])); // GCP Subnets below 50% Utilization

				// Rebuild Powerpoint File(s)
				// ** Using external library to save re-writing the string replacement functions manually. Will probably pull this in as native code at some point.
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Rebuilding Powerpoint Template(s)");
				$injector = new BasicInjector();
				$injector->injectMappingAndCreateNewFile(
					$mapping,
					$SelectedTemplate['ExtractedDir'].'-extracted.pptx',
					$SelectedTemplate['ExtractedDir'].'.pptx'
				);
		
				// Cleanup
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Final Cleanup");
				unlink($SelectedTemplate['ExtractedDir'].'-extracted.pptx');
			}
			// End of new loop

			// Report Status as Done
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Done");
			$this->AssessmentReporting->updateReportEntryStatus($config['UUID'],'Completed');
	
			$Status = 'Success';
		}
	
		## Generate Response
		$response = array(
			'Status' => $Status,
		);
		if (isset($Error)) {
			$response['Error'] = $Error;
		} else {
			$response['id'] = $config['UUID'];
		}
		return $response;
	}
	
	public function createProgress($id,$SelectedTemplates) {
		$Total = count($SelectedTemplates);
		$Templates = array_values(array_column($SelectedTemplates,'FileName'));
		$Progress = json_encode(array(
			'Total' => ($Total * 6) + 9,
			'Count' => 0,
			'Action' => "Starting..",
			'Templates' => $Templates
		));

		// Write the progress file
		if (file_put_contents($this->getDir()['Files'].'/reports/report-'.$id.'.progress', $Progress) === false) {
			die("Unable to save progress file");
		}
	}

	public function writeProgress($id, $count, $action = "") {
		$filePath = $this->getDir()['Files'] . '/reports/report-' . $id . '.progress';
		
		// Read the existing content
		if (file_exists($filePath)) {
			$existingContent = file_get_contents($filePath);
			$progressData = json_decode($existingContent, true);
		} else {
			$progressData = [];
		}
		
		// Update the progress data
		$count++;
		$progressData['Count'] = $count;
		$progressData['Action'] = $action;
		
		// Encode the updated data to JSON
		$newContent = json_encode($progressData);
		
		// Write the updated content back to the file
		if (file_put_contents($filePath, $newContent) === false) {
			die("Unable to save progress file");
		}
		
		return $count;
	}
	
	public function getProgress($id) {
		$ProgressFile = $this->getDir()['Files'].'/reports/report-'.$id.'.progress';
		if (file_exists($ProgressFile)) {
			$myfile = fopen($ProgressFile, "r") or die("0");
			$Current = json_decode(fread($myfile,filesize($ProgressFile)));
			if (isset($Current) && isset($Current->Count) && isset($Current->Action)) {
				return array(
					'Progress' => (100 / $Current->Total) * $Current->Count,
					'Action' => $Current->Action.'..'
				);
			} else {
				return array(
					'Progress' => 0,
					'Action' => 'Checking..'
				);
			}
		} else {
			return array(
				'Progress' => 0,
				'Action' => 'Starting..'
			);
		}
	}
	
	public function getReportFiles() {
		$files = array_diff(scandir($this->getDir()['Files'].'/reports/'),array('.', '..','placeholder.txt'));
		return $files;
	}
}