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
				'TotalAssetsCount' => '{"segments": [],"timeDimensions": [{"dimension": "AssetDetails.doc_updated_at","dateRange": ["'.$StartDimension.'","'.$EndDimension.'"],"granularity": null}],"ungrouped": false,"dimensions": [],"measures": ["AssetDetails.count"], "filters": [{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]}]}',
				'AssetsByClassification' => '{"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails.doc_asset_insight_classification","assetinsighttypes.label","AssetDetails.doc_asset_insight_sub_classification","assetinsightsubclassifications.label"],"filters":[{"member":"AssetDetails.is_valid_sub_classification","operator":"equals","values":["true"]},{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]},{"or":[{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["zombie"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["zombie"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["zombie/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]}]},{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["compliance"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["compliance"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["compliance/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]}]},{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["ghost"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["ghost"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["ghost/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["false"]}]}]}],"timezone":"UTC","segments":[]}',
				'AssetsByClassificationCount' => '{"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":[],"filters":[{"member":"AssetDetails.is_valid_sub_classification","operator":"equals","values":["true"]},{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]},{"or":[{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["zombie"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["zombie"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["zombie/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]}]},{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["compliance"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["compliance"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["compliance/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]}]},{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["ghost"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["ghost"]},{"member":"AssetDetails.doc_asset_insight_state","operator":"equals","values":["ghost/active"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["false"]}]}]}],"timezone":"UTC","segments":[]}',
				'AssetsWithMissingRecords' => '{"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["assetinsightindicators.label","AssetDetails.doc_asset_insight_indicator"],"filters":[{"member":"AssetDetails.doc_asset_insight_indicator","operator":"startsWith","values":["registration"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]},{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]}],"timezone":"UTC","segments":[]}',
				'AssetsWithMissingRecordsCount' => '{"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":[],"filters":[{"member":"AssetDetails.doc_asset_insight_indicator","operator":"startsWith","values":["registration"]},{"member":"AssetDetails.doc_asset_managed","operator":"equals","values":["true"]},{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]}],"timezone":"UTC","segments":[]}',
				'AssetsByProvider' => '{"ungrouped":false,"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails.provider_label"],"segments":[],"filters":[{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]}]}',
				'AssetsByLocation' => '{"ungrouped":false,"measures":["AssetDetails.count"],"dimensions":["AssetDetails.location","AssetDetails.doc_asset_region","AssetDetails.provider_label","AssetDetails.provider_location"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"segments":[],"filters":[{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]},{"member":"AssetDetails.doc_asset_region","operator":"notEquals","values":["unknown"]}],"limit":5}',
				'AssetsByCategory' => '{"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["assetcategories.name","AssetDetails.doc_asset_category"],"filters":[{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]}],"timezone":"UTC","segments":[]}',
				'AssetInsightIndicators' => '{"ungrouped":true,"measures":[],"dimensions":["assetinsightindicators.insightindicator","assetinsightindicators.insightindicator_key","assetinsightindicators.label"],"segments":[]}',
				'AssetLocations' => '{"dimensions": ["assetlocations.provider","assetlocations.provider_label","assetlocations.region","assetlocations.location","assetlocations.provider_location","assetlocations.provider_region"],"measures": [],"segments": [],"ungrouped": true}',
				'ZombieAssetsByIndicator' => '{"ungrouped":false,"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails.doc_asset_insight_indicator"],"segments":[],"filters":[{"member":"AssetDetails.is_valid_sub_classification","operator":"equals","values":["true"]},{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]},{"and":[{"member":"AssetDetails.doc_asset_insight_classification","operator":"equals","values":["zombie"]},{"member":"AssetDetails.doc_asset_insight_sub_classification","operator":"startsWith","values":["zombie"]}]}]}',
				'CloudSubnetUtilizationAbove50' => '{"measures":[],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]},{"operator":"gte","member":"NetworkInsightsSubnet.utilization_percent","values":["50"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.ref_id_resource_id","NetworkInsightsSubnet.address","NetworkInsightsSubnet.utilization_percent","NetworkInsightsSubnet.provider","NetworkInsightsSubnet.usage","NetworkInsightsSubnet.name"],"timeDimensions":[{"dimension":"NetworkInsightsSubnet.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"limit":3}',
				'CloudSubnetUtilizationAbove50Count' => '{"measures":["NetworkInsightsSubnet.count"],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]},{"operator":"gte","member":"NetworkInsightsSubnet.utilization_percent","values":["50"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.provider"],"timeDimensions":[{"dimension":"NetworkInsightsSubnet.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}]}',
				'CloudSubnetUtilizationBelow50' => '{"measures":[],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]},{"operator":"lt","member":"NetworkInsightsSubnet.utilization_percent","values":["50"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.ref_id_resource_id","NetworkInsightsSubnet.address","NetworkInsightsSubnet.utilization_percent","NetworkInsightsSubnet.provider","NetworkInsightsSubnet.usage","NetworkInsightsSubnet.name"],"timeDimensions":[{"dimension":"NetworkInsightsSubnet.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"limit":3}',
				'CloudSubnetUtilizationBelow50Count' => '{"measures":["NetworkInsightsSubnet.count"],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]},{"operator":"lt","member":"NetworkInsightsSubnet.utilization_percent","values":["50"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.provider"],"timeDimensions":[{"dimension":"NetworkInsightsSubnet.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}]}',
				'CloudSubnetsByProvider' => '{"dimensions":["NetworkInsightsSubnet.provider"],"ungrouped":false,"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"dimension":"NetworkInsightsSubnet.updated_at","granularity":null}],"measures":["NetworkInsightsSubnet.count"],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]}]}',
				// 'CloudIPsByProvider' => '{"measures":["NetworkInsightsSubnet.count"],"segments":[],"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"dimension":"NetworkInsightsSubnet.updated_at","granularity":null}],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.provider","values":["AWS","GCP","Azure"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.utilization_used","NetworkInsightsSubnet.utilization_total","NetworkInsightsSubnet.provider"]}',
				'CloudIPsByProvider' => '{"ungrouped":false,"timeDimensions": [{"dimension": "AssetDetails.doc_updated_at","dateRange": ["'.$StartDimension.'","'.$EndDimension.'"],"granularity": null}],"measures":["AssetDetails.count"],"dimensions":["AssetDetails.provider_label"],"segments":[],"filters":[{"and":[{"member":"AssetDetails.doc_asset_ip_address","operator":"set"},{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]}]}]}',
				'CloudDNSZonesByProvider' => '{"ungrouped":false,"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails.provider_label","AssetDetails.doc_asset_category"],"segments":[],"filters":[{"and":[{"member":"AssetDetails.doc_asset_category","operator":"equals","values":["dns"]},{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]}]}]}',
				'HighRiskDNSRecords' => '{"measures":["NetworkInsightsDnsRecords.count"],"timeDimensions":[{"dimension":"NetworkInsightsDnsRecords.evaluation_time","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["NetworkInsightsDnsRecords.indicator_id"],"filters":[{"member":"NetworkInsightsDnsRecords.indicator_id","operator":"set"},{"member":"NetworkInsightsDnsRecords.provider","operator":"equals","values":["amazon_web_service","microsoft_azure","google_cloud_platform"]}],"timezone":"UTC","segments":[]}',
				'CloudDNSRecordsByProvider' => '{"measures":["NetworkInsightsDnsRecords.count"],"timeDimensions":[{"dimension":"NetworkInsightsDnsRecords.evaluation_time","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"segments":[],"ungrouped":false,"dimensions":["NetworkInsightsDnsRecords.provider"],"filters":[{"member":"NetworkInsightsDnsRecords.provider","operator":"equals","values":["amazon_web_service","microsoft_azure","google_cloud_platform"]}]}',
				// Original Value - Summary number of overlapping subnets
				//'CloudSubnetOverlapCount' => '{"measures":["NetworkInsightsOverlappingBlocksList.count_total"],"timeDimensions":[{"dimension":"NetworkInsightsOverlappingBlocksList.generated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":[],"filters":[],"timezone":"UTC","segments":[]}',
				'LicensingManagement' => '{"measures":["TokenUtilManagementObjects.count"],"timeDimensions":[{"dimension":"TokenUtilManagementObjects.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":"day"}],"dimensions":["TokenUtilManagementObjects.category","TokenUtilManagementObjects.object_type"],"filters":[{"and":[{"member":"TokenUtilManagementObjects.object_type","operator":"equals","values":["DDI","Active IPs","Assets"]},{"member":"TokenUtilManagementObjects.category","operator":"equals","values":["Native"]}]}]}',
				'LicensingServer' => '{"measures":["TokenUtilProtoSrvSM.count","TokenUtilProtoSrvSM.tokens"],"timeDimensions":[{"dimension":"TokenUtilProtoSrvSM.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":"day"}],"dimensions":[]}',
				'LicensingReporting' => '{"measures":["TokenUtilReporting.count","TokenUtilReporting.tokens"],"timeDimensions":[{"dimension":"TokenUtilReporting.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":"day"}],"dimensions":[]}',
				'OverlappingSubnets' => '{"dimensions":["NetworkInsightsOverlappingBlocksRolled.address","NetworkInsightsOverlappingBlocksRolled.overlap_count","NetworkInsightsOverlappingBlocksRolled.sources","NetworkInsightsOverlappingBlocksRolled.providers"],"filters":[{"member":"NetworkInsightsOverlappingBlocksRolled.sources_str","operator":"contains","values":["Cloud"]}],"timeDimensions":[],"limit":3,"offset":0,"total":true,"order":{"NetworkInsightsOverlappingBlocksRolled.overlap_count":"desc"}}',
				'OverlappingZones' => '{"ungrouped":false,"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails.provider_label","AssetDetails.doc_asset_category","AssetDetails.doc_name"],"segments":[],"filters":[{"and":[{"member":"AssetDetails.doc_asset_category","operator":"equals","values":["dns"]},{"member":"AssetDetails.provider_label","operator":"equals","values":["AWS","Azure","GCP"]}]}]}',
				'AbandonedRecords' => '{"dimensions":["NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type","NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str"],"filters":[{"member":"NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids_str","operator":"contains","values":["Abandoned"]}],"timeDimensions":[{"dimension":"NetworkInsightsDnsRecordsRolledByIndicatorId.evaluation_time","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null}],"limit":3,"offset":0,"ungrouped":true,"total":true}',
				'UntrustedRecords' => '{"dimensions":["NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type","NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str"],"filters":[{"member":"NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids_str","operator":"contains","values":["Untrusted"]}],"timeDimensions":[{"dimension":"NetworkInsightsDnsRecordsRolledByIndicatorId.evaluation_time","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null}],"limit":3,"offset":0,"ungrouped":true,"total":true}',
				'DanglingRecords' => '{"dimensions":["NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type","NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str"],"filters":[{"member":"NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids_str","operator":"contains","values":["Dangling"]}],"timeDimensions":[{"dimension":"NetworkInsightsDnsRecordsRolledByIndicatorId.evaluation_time","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null}],"limit":3,"offset":0,"ungrouped":true,"total":true}'
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
				'AssetsByProvider' => 0, // Slide 7
				'AssetsByLocation' => 1, // Slide 7
				'ZombieAssetsByIndicator' => array(
					'Landscape' => 3,
					'Portrait' => 2
				), // Slide 8
				'AssetsByCategory' => array(
					'Landscape' => 2,
					'Portrait' => 3
				), // Slide 8
				'ZombieAssetsTable' => 4, // Slide 8
				'AssetsWithMissingRecords' => 5, // Slide 9
				'NonCompliantAssets' => 6, // Slide 9
				'AssetsWithMissingRecordsTable' => 7, // Slide 9
				'TotalIPsByProvider' => 8, // Slide 11
				'OverlappingSubnetsTable' => 9, // Slide 11
				'OverutilizedSubnetsTable' => 10, // Slide 13
				'UnderutilizedSubnetsTable' => 11, // Slide 13
				'HighRiskDNSRecordsByCategory' => 12, // Slide 16
				'DanglingRecordsTable' => 13, // Slide 16
				'AbandonedRecordsTable' => 14, // Slide 17
				'UntrustedRecordsTable' => 15 // Slide 17
			];

			// Function to get the full path of the file based on the sheet name
			function getEmbeddedSheetFilePath($sheetName, $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $orientation) {
				if (isset($EmbeddedSheets[$sheetName])) {
					$fileIndex = $EmbeddedSheets[$sheetName];
					if (is_array($fileIndex)) {
						// Check for orientation
						if (isset($fileIndex[$orientation])) {
							$fileIndex = $fileIndex[$orientation];
						}
					} else {
						$fileIndex = $fileIndex;
					}
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
			$NonCompliantAssets = array();
			$ZombieIdleAssetsCount = 0;
			$ZombieOrphanedAssetsCount = 0;
			$ZombieAssetsResourceUtilizationCount = 0;
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
						$NonCompliantAssets[$value->{'assetinsightsubclassifications.label'}] = $value->{'AssetDetails.count'};
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
			$CloudSubnetUtilizationAbove50Count = $CubeJSResults['CloudSubnetUtilizationAbove50Count']['Body']->result->data ?? array();
			$CloudSubnetUtilizationAbove50ByProvider = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Total" => 0];
			$CloudSubnetUtilizationAbove50Total = 0;
			foreach ($CloudSubnetUtilizationAbove50Count as $value) {
				$CloudSubnetUtilizationAbove50ByProvider[$value->{'NetworkInsightsSubnet.provider'}] = $value->{'NetworkInsightsSubnet.count'} ?? 0;
				$CloudSubnetUtilizationAbove50Total += $value->{'NetworkInsightsSubnet.count'} ?? 0;
			}
			$CloudSubnetUtilizationBelow50Count = $CubeJSResults['CloudSubnetUtilizationBelow50Count']['Body']->result->data ?? array();
			$CloudSubnetUtilizationBelow50ByProvider = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Total" => 0];
			$CloudSubnetUtilizationBelow50Total = 0;
			foreach ($CloudSubnetUtilizationBelow50Count as $value) {
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

			$AWSIPsCount = 0;
			$AzureIPsCount = 0;
			$GCPIPsCount = 0;
			$TotalCloudIPsCount = 0;
			foreach ($CloudIPsByProvider as $value) {
				switch ($value->{'AssetDetails.provider_label'}) {
					case 'AWS':
						$AWSIPsCount += $value->{'AssetDetails.count'} ?? 0;
						$TotalCloudIPsCount += $value->{'AssetDetails.count'} ?? 0;
						break;
					case 'Azure':
						$AzureIPsCount += $value->{'AssetDetails.count'} ?? 0;
						$TotalCloudIPsCount += $value->{'AssetDetails.count'} ?? 0;
						break;
					case 'GCP':
						$GCPIPsCount += $value->{'AssetDetails.count'} ?? 0;
						$TotalCloudIPsCount += $value->{'AssetDetails.count'} ?? 0;
						break;
				}
				$TotalSubnetsCount += $value->{'AssetDetails.count'} ?? 0;
			}

			// $CloudIPsByProviderTotal = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Total" => 0];
			// $CloudIPsByProviderTotalUsed = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Total" => 0];
			// foreach ($CloudIPsByProvider as $CloudIPsByProviderEntry) {
			// 	$Provider = $CloudIPsByProviderEntry->{'NetworkInsightsSubnet.provider'};
			// 	$CloudIPsByProviderTotal[$Provider] += (int)$CloudIPsByProviderEntry->{'NetworkInsightsSubnet.utilization_total'};
			// 	$CloudIPsByProviderTotal["Total"] += (int)$CloudIPsByProviderEntry->{'NetworkInsightsSubnet.utilization_total'};
			// 	$CloudIPsByProviderTotalUsed[$Provider] += (int)$CloudIPsByProviderEntry->{'NetworkInsightsSubnet.utilization_used'};
			// 	$CloudIPsByProviderTotalUsed["Total"] += (int)$CloudIPsByProviderEntry->{'NetworkInsightsSubnet.utilization_used'};
			// }

			// DNS Zones by Provider
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building DNS Zones by Provider");
			$CloudDNSZonesByProvider = $CubeJSResults['CloudDNSZonesByProvider']['Body']->result->data ?? array();
			$CloudDNSZonesByProviderTotals = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Total" => 0];
			foreach ($CloudDNSZonesByProvider as $CloudDNSZonesByProviderEntry) {
				$Provider = $CloudDNSZonesByProviderEntry->{'AssetDetails.provider_label'};
				$CloudDNSZonesByProviderTotals[$Provider] += (int)$CloudDNSZonesByProviderEntry->{'AssetDetails.count'};
				$CloudDNSZonesByProviderTotals["Total"] += (int)$CloudDNSZonesByProviderEntry->{'AssetDetails.count'};
			}


			// High Risk DNS Records
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building High Risk DNS Records");
			$HighRiskDNSRecords = $CubeJSResults['HighRiskDNSRecords']['Body']->result->data ?? array();
			$AbandonedDNSCount = 0;
			$UntrustedDNSCount = 0;
			$DanglingDNSCount = 0;
			$HighRiskDNSRecordsCount = 0;
			
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
				$HighRiskDNSRecordsCount += $record->{'NetworkInsightsDnsRecords.count'};
			}
			
			// High Risk DNS Records
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building DNS Records by Provider");
			$CloudDNSRecordsByProvider = $CubeJSResults['CloudDNSRecordsByProvider']['Body']->result->data ?? array();
			$CloudDNSRecordsByProviderTotals = ["amazon_web_service" => 0, "microsoft_azure" => 0, "google_cloud_platform" => 0, "Total" => 0];
			
			foreach ($CloudDNSRecordsByProvider as $CloudDNSRecordsByProviderEntry) {
				$Provider = $CloudDNSRecordsByProviderEntry->{'NetworkInsightsDnsRecords.provider'};
				$CloudDNSRecordsByProviderTotals[$Provider] += (int)$CloudDNSRecordsByProviderEntry->{'NetworkInsightsDnsRecords.count'};
				$CloudDNSRecordsByProviderTotals["Total"] += (int)$CloudDNSRecordsByProviderEntry->{'NetworkInsightsDnsRecords.count'};
			}

			// Subnet Overlap Count
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Overlapping Subnets");

			// Original Value - Summary number of overlapping subnets
			//$CloudSubnetOverlapCount = $CubeJSResults['CloudSubnetOverlapCount']['Body']->result->data[0]->{'NetworkInsightsOverlappingBlocksList.count_total'} ?? 0;

			// New Value - Total number of overlapping subnets
			$CloudSubnetOverlapCount = 0;
			foreach ($CubeJSResults['OverlappingSubnets']['Body']->result->data as $Overlap) {
				$CloudSubnetOverlapCount += $Overlap->{'NetworkInsightsOverlappingBlocksRolled.overlap_count'} ?? 0;
			}

			// Licensing
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building License Utilization");
			$ActiveIPCount = 0;
			$AssetCount = 0;
			$DDIObjectCount = 0;
			$TokensManagement = 0;
			$TokensActiveIPCount = 0;
			$TokensActiveIPPercentage = 0;
			$TokensAssetCount = 0;
			$TokensAssetPercentage = 0;
			$TokensDDICount = 0;
			$TokensDDIPercentage = 0;
			$TokensServer = 0;
			$TokensReporting = 0;

			$LicensingManagement = $CubeJSResults['LicensingManagement']['Body']->result->data ?? array();
			foreach ($LicensingManagement as $value) {
				switch ($value->{'TokenUtilManagementObjects.object_type'}) {
					case 'Active IPs':
						$ActiveIPCount += $value->{'TokenUtilManagementObjects.count'};
						break;
					case 'Assets':
						$AssetCount += $value->{'TokenUtilManagementObjects.count'};
						break;
					case 'DDI':
						$DDIObjectCount += $value->{'TokenUtilManagementObjects.count'};
						break;
				}
			}

			$TokensActiveIPCount = $ActiveIPCount / 13;
			$TokensAssetCount = $AssetCount / 3;
			$TokensDDICount = $DDIObjectCount / 25;
			$TokensManagement = ($TokensActiveIPCount+$TokensAssetCount+$TokensDDICount);
			$TokensManagementCount = $TokensManagement*1.2; // Token Count + 20%
			$TokensActiveIPPercentage = (100 / $TokensManagement) * $TokensActiveIPCount;
			$TokensAssetPercentage = (100 / $TokensManagement) * $TokensAssetCount;
			$TokensDDIPercentage = (100 / $TokensManagement) * $TokensDDICount;

			$TokensServer = $CubeJSResults['LicensingServer']['Body']->result->data[0]->{'TokenUtilProtoSrvSM.tokens'} ?? 0;
			if ($TokensServer != 0) {
				$TokensServer = $TokensServer*1.2; // Token Count + 20%
			} else {
				$TokensServer = 0;
			}
			$TokensReporting = $CubeJSResults['LicensingReporting']['Body']->result->data[0]->{'TokenUtilReporting.tokens'} ?? 0;
			if ($TokensReporting != 0) {
				$TokensReporting = $TokensReporting*1.2; // Token Count + 20%
			} else {
				$TokensReporting = 0;
			}

			// Top 3 Zombie Assets - Slide 8
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Top 3 Zombie Assets");
			$JSONData = json_decode('{"_base_filters":[{"application":"cloud_discovery"},{"application":"asset_discovery"}],"limit":"3","offset":null,"query":"(asset_insight_classification == \'zombie\') AND (asset_context == \'cloud\')","filters":[{"application":"cloud_discovery"},{"application":"asset_discovery"}],"field_filters":[{"key":"asset_managed","values":["true"]},{"key":"updated_at","gte":"'.$StartDimension.'","lte":"'.$EndDimension.'"}]}');
			$Top3ZombieAssetsResponse = $this->QueryCSP("post","atlas-search-api/v1/discover",$JSONData);
			$Top3ZombieAssets = $Top3ZombieAssetsResponse->hits->hits ?? array();

			// Top 3 Assets with Missing Records - Slide 9
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Top 3 Assets with Missing Records");
			$JSONData = json_decode('{"_base_filters":[{"application":"cloud_discovery"},{"application":"asset_discovery"}],"limit":"3","offset":null,"query":"(asset_insight_registration_status == \'Unregistered\') AND (asset_context == \'cloud\')","filters":[{"application":"cloud_discovery"},{"application":"asset_discovery"}],"field_filters":[{"key":"asset_managed","values":["true"]},{"key":"updated_at","gte":"'.$StartDimension.'","lte":"'.$EndDimension.'"}],"sort":[{"fields":{"doc.updated_at":{"order":"asc"}}}]}');
			$Top3AssetsWithMissingRecordsResponse = $this->QueryCSP("post","atlas-search-api/v1/discover",$JSONData);
			$Top3AssetsWithMissingRecords = $Top3AssetsWithMissingRecordsResponse->hits->hits ?? array();

			// Cloud Forwarders
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Cloud Forwarders");
			$CloudForwardersResult = $this->QueryCSP("get","api/dns/v1/cloud_forwarder/endpoint");
			$CloudForwarders = $CloudForwardersResult->results ?? array();
			$CloudForwarderCounts = [
				'Microsoft Azure' => ['inbound' => 0, 'outbound' => 0],
				'Amazon Web Services' => ['inbound' => 0, 'outbound' => 0],
				'Google Cloud Platform' => ['inbound' => 0, 'outbound' => 0],
				'Total' => 0
			];
				
			foreach ($CloudForwarders as $CloudForwarder) {
				$provider = $CloudForwarder->provider_type;
				$direction = $CloudForwarder->direction;
				$CloudForwarderCounts[$provider][$direction]++;
				$CloudForwarderCounts['Total']++;
			}

			// Overlapping DNS Zones
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Overlapping DNS Zones");
			$OverlappingZones = $CubeJSResults['OverlappingZones']['Body']->result->data ?? array();
			$OverlappingZonesCount = 0;
			foreach ($OverlappingZones as $OverlappingZone) {
				$OverlappingZoneCount = $OverlappingZone->{'AssetDetails.count'} ?? 0;
				if ($OverlappingZoneCount > 1) {
					$OverlappingZonesCount += $OverlappingZoneCount;
				}
			}

			// DNS Complexity Score
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building DNS Complexity Score");
			$DNSComplexityScore = 0;

			/* Calculate number of Providers */
			$providers = ['AWS', 'Azure', 'GCP'];
			$ProviderCount = count(array_filter($providers, fn($provider) => $CloudDNSZonesByProviderTotals[$provider] > 0));

			/* If provider count is greater than 0, add to score */
			if ($ProviderCount > 0) {
				$DNSComplexityScore += 20;
				if ($ProviderCount > 1) {
					$DNSComplexityScore += ($ProviderCount - 1) * 50;
				}
			}

			/* Add points per number of records using these thresholds */
			$dnsThresholds = [100 => 10, 1000 => 50, PHP_INT_MAX => 75];
			foreach (['amazon_web_service', 'microsoft_azure', 'google_cloud_platform'] as $provider) {
				$total = $CloudDNSRecordsByProviderTotals[$provider];
				foreach ($dnsThresholds as $threshold => $score) {
					if ($total > $threshold) {
						$DNSComplexityScore += $score;
					}
				}
			}

			/* Add points per number of inbound and outbound forwarders */
			$DNSComplexityScore += $CloudForwarderCounts['Total'] * 25;

			/* Add points per number of overlapping zones */
			$DNSComplexityScore += $OverlappingZonesCount * 100;

			// Loop for each selected template
			foreach ($SelectedTemplates as &$SelectedTemplate) {
				$embeddedDirectory = $SelectedTemplate['ExtractedDir'].'/ppt/embeddings/';
				$embeddedFiles = array_values(array_diff(scandir($embeddedDirectory), array('.', '..')));
				usort($embeddedFiles, 'strnatcmp');
				$this->logging->writeLog("Assessment","Embedded Files List","debug",['Template' => $SelectedTemplate, 'Embedded Files' => $embeddedFiles]);


				// ** CHARTS & TABLES ** //
				// Assets By Provider - Slide 7
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Assets by Provider");
				$AssetsByProvider = $CubeJSResults['AssetsByProvider']['Body'];
				if (isset($AssetsByProvider->result->data)) {
					$EmbeddedAssetsByProvider = getEmbeddedSheetFilePath('AssetsByProvider', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$AssetsByProviderSS = IOFactory::load($EmbeddedAssetsByProvider);
					$RowNo = 2;
					foreach ($AssetsByProvider->result->data as $ProviderType) {
						$AssetsByProviderS = $AssetsByProviderSS->getActiveSheet();
						$AssetsByProviderS->setCellValue('A'.$RowNo, $ProviderType->{'AssetDetails.provider_label'});
						$AssetsByProviderS->setCellValue('B'.$RowNo, $ProviderType->{'AssetDetails.count'});
						$RowNo++;
					}
					$AssetsByProviderW = IOFactory::createWriter($AssetsByProviderSS, 'Xlsx');
					$AssetsByProviderW->save($EmbeddedAssetsByProvider);
				}

				// Assets By Location - Slide 7
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Assets by Location");
				$AssetsByLocation = $CubeJSResults['AssetsByLocation']['Body'];
				if (isset($AssetsByLocation->result->data)) {
					$EmbeddedAssetsByLocation = getEmbeddedSheetFilePath('AssetsByLocation', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$AssetsByLocationSS = IOFactory::load($EmbeddedAssetsByLocation);
					$RowNo = 2;
					$AssetsByLocationIncludedCount = 0;
					foreach ($AssetsByLocation->result->data as $AssetLocation) {
						$AssetsByLocationS = $AssetsByLocationSS->getActiveSheet();
						$AssetsByLocationS->setCellValue('A'.$RowNo, $AssetLocation->{'AssetDetails.provider_location'});
						$AssetsByLocationS->setCellValue('B'.$RowNo, round((100 / $TotalAssetsCount) * $AssetLocation->{'AssetDetails.count'},2) / 100);
						$AssetsByLocationIncludedCount += $AssetLocation->{'AssetDetails.count'};
						$RowNo++;
					}
					$AssetsByLocationExcludedCount = $TotalAssetsCount - $AssetsByLocationIncludedCount;
					if ($AssetsByLocationExcludedCount > 0) {
						$AssetsByLocationS = $AssetsByLocationSS->getActiveSheet();
						$AssetsByLocationS->setCellValue('A'.$RowNo, 'Others');
						$AssetsByLocationS->setCellValue('B'.$RowNo, round((100 / $TotalAssetsCount) * $AssetsByLocationExcludedCount,2) / 100);
					}
					$AssetsByLocationW = IOFactory::createWriter($AssetsByLocationSS, 'Xlsx');
					$AssetsByLocationW->save($EmbeddedAssetsByLocation);
				}

				// Assets By Category - Slide 8
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Assets by Category");
				$AssetsByCategory = $CubeJSResults['AssetsByCategory']['Body'];
				if (isset($AssetsByCategory->result->data)) {
					$EmbeddedAssetsByCategory = getEmbeddedSheetFilePath('AssetsByCategory', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$AssetsByCategorySS = IOFactory::load($EmbeddedAssetsByCategory);
					$RowNo = 2;
					foreach ($AssetsByCategory->result->data as $AssetCategory) {
						$AssetsByCategoryS = $AssetsByCategorySS->getActiveSheet();
						$AssetsByCategoryS->setCellValue('A'.$RowNo, $AssetCategory->{'assetcategories.name'});
						$AssetsByCategoryS->setCellValue('B'.$RowNo, $AssetCategory->{'AssetDetails.count'});
						$RowNo++;
					}
					$AssetsByCategoryW = IOFactory::createWriter($AssetsByCategorySS, 'Xlsx');
					$AssetsByCategoryW->save($EmbeddedAssetsByCategory);
				}

				// Zombie Assets by Classification - Slide 8
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Zombie Assets by Indicator");
				$AssetInsightIndicators = $CubeJSResults['AssetInsightIndicators']['Body']->result->data ?? array();
				$AssetInsightIndicatorsKeys = array_column($AssetInsightIndicators, 'assetinsightindicators.insightindicator_key');
				$AssetInsightIndicatorsLabels = array_column($AssetInsightIndicators, 'assetinsightindicators.label');
				$ZombieAssetsByIndicator = $CubeJSResults['ZombieAssetsByIndicator']['Body'];

				if (isset($ZombieAssetsByIndicator->result->data)) {
					$EmbeddedZombieAssetsByIndicator = getEmbeddedSheetFilePath('ZombieAssetsByIndicator', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$ZombieAssetsByIndicatorSS = IOFactory::load($EmbeddedZombieAssetsByIndicator);
					$RowNo = 2;
					foreach ($ZombieAssetsByIndicator->result->data as $ZombieAssetIndicator) {
						$ZombieAssetsByIndicatorS = $ZombieAssetsByIndicatorSS->getActiveSheet();

						$Indicator = implode(',',array_slice(explode('/',$ZombieAssetIndicator->{'AssetDetails.doc_asset_insight_indicator'}),1));
						$IndicatorIndex = array_search($Indicator, $AssetInsightIndicatorsKeys);
						$IndicatorLabel = $IndicatorIndex !== false ? $AssetInsightIndicatorsLabels[$IndicatorIndex] : null;

						$ZombieAssetsByIndicatorS->setCellValue('A'.$RowNo, $IndicatorLabel);
						$ZombieAssetsByIndicatorS->setCellValue('B'.$RowNo, $ZombieAssetIndicator->{'AssetDetails.count'});
						$RowNo++;
					}
					$ZombieAssetsByIndicatorW = IOFactory::createWriter($ZombieAssetsByIndicatorSS, 'Xlsx');
					$ZombieAssetsByIndicatorW->save($EmbeddedZombieAssetsByIndicator);
				}

				// Asset Regions
				$AssetLocations = $CubeJSResults['AssetLocations']['Body']->result->data ?? array();
				$AssetLocationsRegionKeys = array_column($AssetLocations, 'assetlocations.region');
				$AssetLocationsLocationKeys = array_column($AssetLocations, 'assetlocations.location');

				// Zombie Assets Table - Slide 8
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Zombie Assets Table");
				if (isset($Top3ZombieAssets)) {
					$EmbeddedZombieAssetsTable = getEmbeddedSheetFilePath('ZombieAssetsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$ZombieAssetsTableSS = IOFactory::load($EmbeddedZombieAssetsTable);
					$RowNo = 2;
					foreach ($Top3ZombieAssets as $Top3ZombieAsset) {
						$ZombieAssetsTableS = $ZombieAssetsTableSS->getActiveSheet();
						$ZombieAssetsTableS->setCellValue('A'.$RowNo, $Top3ZombieAsset->{'_source'}->{'doc'}->{'name'}); // Asset Name
						$ZombieAssetsTableS->setCellValue('B'.$RowNo, $Top3ZombieAsset->{'_source'}->{'doc'}->{'asset_vendor'}); // Asset Vendor

						$RegionIndex = array_search($Top3ZombieAsset->{'_source'}->{'doc'}->{'asset_region'}, $AssetLocationsRegionKeys);
						$RegionLocation = $RegionIndex !== false ? $AssetLocationsLocationKeys[$RegionIndex] : null;

						$ZombieAssetsTableS->setCellValue('C'.$RowNo, $RegionLocation); // Asset Location (Region)
						$ZombieAssetsTableS->setCellValue('D'.$RowNo, $Top3ZombieAsset->{'_source'}->{'doc'}->{'asset_insight_classification'}[0]); // Asset Classification
						$ZombieAssetsTableS->setCellValue('E'.$RowNo, $Top3ZombieAsset->{'_source'}->{'doc'}->{'asset_ip_address'}[0] ?? null); // Asset IP Address
						$ZombieAssetsTableS->setCellValue('F'.$RowNo, implode(',',array_slice(explode('/',$Top3ZombieAsset->{'_source'}->{'doc'}->{'asset_insight_sub_classification'}[0]),1))); // Asset Sub-Classification
						$ZombieAssetsTableS->setCellValue('G'.$RowNo, $this->timeAgo($Top3ZombieAsset->{'_source'}->{'doc'}->{'updated_at'})); // Asset Last Seen
						$ZombieAssetsTableS->setCellValue('H'.$RowNo, $Top3ZombieAsset->{'_source'}->{'doc'}->{'asset_provider'}); // Asset Provider
						$RowNo++;
					}
					$ZombieAssetsTableW = IOFactory::createWriter($ZombieAssetsTableSS, 'Xlsx');
					$ZombieAssetsTableW->save($EmbeddedZombieAssetsTable);
				}

				// Assets with Missing Records - Slide 9
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Assets with Missing Records");
				$AssetsWithMissingRecords = $CubeJSResults['AssetsWithMissingRecords']['Body'];
				if (isset($AssetsWithMissingRecords->result->data)) {
					$EmbeddedAssetsWithMissingRecords = getEmbeddedSheetFilePath('AssetsWithMissingRecords', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$AssetsWithMissingRecordsSS = IOFactory::load($EmbeddedAssetsWithMissingRecords);
					$RowNo = 2;
					foreach ($AssetsWithMissingRecords->result->data as $ProviderType) {
						$AssetsWithMissingRecordsS = $AssetsWithMissingRecordsSS->getActiveSheet();
						$AssetsWithMissingRecordsS->setCellValue('A'.$RowNo, $ProviderType->{'assetinsightindicators.label'});
						$AssetsWithMissingRecordsS->setCellValue('B'.$RowNo, $ProviderType->{'AssetDetails.count'});
						$RowNo++;
					}
					$AssetsWithMissingRecordsW = IOFactory::createWriter($AssetsWithMissingRecordsSS, 'Xlsx');
					$AssetsWithMissingRecordsW->save($EmbeddedAssetsWithMissingRecords);
				}

				// Assets with Missing Records Table - Slide 9
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Assets with Missing Records Table");
				if (isset($Top3ZombieAssets)) {
					$EmbeddedAssetsWithMissingRecordsTable = getEmbeddedSheetFilePath('AssetsWithMissingRecordsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$AssetsWithMissingRecordsTableSS = IOFactory::load($EmbeddedAssetsWithMissingRecordsTable);
					$RowNo = 2;
					foreach ($Top3AssetsWithMissingRecords as $Top3AssetWithMissingRecords) {
						$AssetsWithMissingRecordsTableS = $AssetsWithMissingRecordsTableSS->getActiveSheet();
						$AssetsWithMissingRecordsTableS->setCellValue('A'.$RowNo, $Top3AssetWithMissingRecords->{'_source'}->{'doc'}->{'name'}); // Asset Name
						$AssetsWithMissingRecordsTableS->setCellValue('B'.$RowNo, $Top3AssetWithMissingRecords->{'_source'}->{'doc'}->{'asset_vendor'}); // Asset Vendor

						$RegionIndex = array_search($Top3AssetWithMissingRecords->{'_source'}->{'doc'}->{'asset_region'}, $AssetLocationsRegionKeys);
						$RegionLocation = $RegionIndex !== false ? $AssetLocationsLocationKeys[$RegionIndex] : null;

						$AssetsWithMissingRecordsTableS->setCellValue('C'.$RowNo, $RegionLocation); // Asset Location (Region)
						$AssetsWithMissingRecordsTableS->setCellValue('D'.$RowNo, $Top3AssetWithMissingRecords->{'_source'}->{'doc'}->{'asset_insight_classification'}[0]); // Asset Classification
						$AssetsWithMissingRecordsTableS->setCellValue('E'.$RowNo, $Top3AssetWithMissingRecords->{'_source'}->{'doc'}->{'asset_ip_address'}[0] ?? null); // Asset IP Address
						$AssetsWithMissingRecordsTableS->setCellValue('F'.$RowNo, $Top3AssetWithMissingRecords->{'_source'}->{'doc'}->{'asset_insight_registration_indicator'}[0]); // Asset Registration Indicator
						$AssetsWithMissingRecordsTableS->setCellValue('G'.$RowNo, $this->timeAgo($Top3AssetWithMissingRecords->{'_source'}->{'doc'}->{'updated_at'})); // Asset Last Seen
						$AssetsWithMissingRecordsTableS->setCellValue('H'.$RowNo, $Top3AssetWithMissingRecords->{'_source'}->{'doc'}->{'asset_provider'}); // Asset Provider
						$RowNo++;
					}
					$AssetsWithMissingRecordsTableW = IOFactory::createWriter($AssetsWithMissingRecordsTableSS, 'Xlsx');
					$AssetsWithMissingRecordsTableW->save($EmbeddedAssetsWithMissingRecordsTable);
				}


				// Non-Compliant Assets - Slide 9
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Non-Compliant Assets");
				if (!empty($NonCompliantAssets)) {
					$EmbeddedNonCompliantAssets = getEmbeddedSheetFilePath('NonCompliantAssets', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$NonCompliantAssetsSS = IOFactory::load($EmbeddedNonCompliantAssets);
					$RowNo = 2;
					foreach ($NonCompliantAssets as $InsightType => $Count) {
						$NonCompliantAssetsS = $NonCompliantAssetsSS->getActiveSheet();
						$NonCompliantAssetsS->setCellValue('A'.$RowNo, $InsightType);
						$NonCompliantAssetsS->setCellValue('B'.$RowNo, $Count);
						$RowNo++;
					}
					$NonCompliantAssetsW = IOFactory::createWriter($NonCompliantAssetsSS, 'Xlsx');
					$NonCompliantAssetsW->save($EmbeddedNonCompliantAssets);
				}

				// Overlapping Subnets - Slide 11
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Overlapping Subnets");
				$OverlappingSubnets = $CubeJSResults['OverlappingSubnets']['Body'];
				if (isset($OverlappingSubnets->result->data)) {
					$EmbeddedOverlappingSubnets = getEmbeddedSheetFilePath('OverlappingSubnetsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$OverlappingSubnetsSS = IOFactory::load($EmbeddedOverlappingSubnets);
					$RowNo = 2;
					foreach ($OverlappingSubnets->result->data as $ProviderType) {
						$OverlappingSubnetsS = $OverlappingSubnetsSS->getActiveSheet();
						$OverlappingSubnetsS->setCellValue('A'.$RowNo, $ProviderType->{'NetworkInsightsOverlappingBlocksRolled.address'});
						$OverlappingSubnetsS->setCellValue('B'.$RowNo, $ProviderType->{'NetworkInsightsOverlappingBlocksRolled.overlap_count'});
						$OverlappingSubnetsS->setCellValue('C'.$RowNo, implode(', ',$ProviderType->{'NetworkInsightsOverlappingBlocksRolled.providers'}));
						$RowNo++;
					}
					$OverlappingSubnetsW = IOFactory::createWriter($OverlappingSubnetsSS, 'Xlsx');
					$OverlappingSubnetsW->save($EmbeddedOverlappingSubnets);
				}

				// Total IPs By Provider - Slide 11
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Total IPs by Provider");
				$EmbeddedTotalIPsByProvider = getEmbeddedSheetFilePath('TotalIPsByProvider', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$TotalIPsByProviderSS = IOFactory::load($EmbeddedTotalIPsByProvider);
				$TotalIPsByProviderS = $TotalIPsByProviderSS->getActiveSheet();
				$RowNo = 2;
				if ($GCPIPsCount > 0) {
					$TotalIPsByProviderS->setCellValue('A'.$RowNo, 'GCP');
					$TotalIPsByProviderS->setCellValue('B'.$RowNo, $GCPIPsCount);
					$RowNo++;
				}
				if ($AzureIPsCount > 0) {
					$TotalIPsByProviderS->setCellValue('A'.$RowNo, 'Azure');
					$TotalIPsByProviderS->setCellValue('B'.$RowNo, $AzureIPsCount);
					$RowNo++;
				}
				if ($AWSIPsCount > 0) {
					$TotalIPsByProviderS->setCellValue('A'.$RowNo, 'AWS');
					$TotalIPsByProviderS->setCellValue('B'.$RowNo, $AWSIPsCount);
					$RowNo++;
				}
				$TotalIPsByProviderW = IOFactory::createWriter($TotalIPsByProviderSS, 'Xlsx');
				$TotalIPsByProviderW->save($EmbeddedTotalIPsByProvider);

				// Overutilized Subnets - Slide 13
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Overutilized Subnets");
				$OverutilizedSubnets = $CubeJSResults['CloudSubnetUtilizationAbove50']['Body'];
				if (isset($OverutilizedSubnets->result->data)) {
					$EmbeddedOverutilizedSubnets = getEmbeddedSheetFilePath('OverutilizedSubnetsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$OverutilizedSubnetsSS = IOFactory::load($EmbeddedOverutilizedSubnets);
					$RowNo = 2;
					foreach ($OverutilizedSubnets->result->data as $ProviderType) {
						$OverutilizedSubnetsS = $OverutilizedSubnetsSS->getActiveSheet();
						$OverutilizedSubnetsS->setCellValue('A'.$RowNo, $ProviderType->{'NetworkInsightsSubnet.address'});
						$OverutilizedSubnetsS->setCellValue('B'.$RowNo, $ProviderType->{'NetworkInsightsSubnet.name'});
						$OverutilizedSubnetsS->setCellValue('C'.$RowNo, $ProviderType->{'NetworkInsightsSubnet.provider'});
						$OverutilizedSubnetsS->setCellValue('D'.$RowNo, round($ProviderType->{'NetworkInsightsSubnet.utilization_percent'},1) . '%');
						$RowNo++;
					}
					$OverutilizedSubnetsW = IOFactory::createWriter($OverutilizedSubnetsSS, 'Xlsx');
					$OverutilizedSubnetsW->save($EmbeddedOverutilizedSubnets);
				}
				
				// Underutilized Subnets - Slide 13
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Underutilized Subnets");
				$UnderutilizedSubnets = $CubeJSResults['CloudSubnetUtilizationBelow50']['Body'];
				if (isset($UnderutilizedSubnets->result->data)) {
					$EmbeddedUnderutilizedSubnets = getEmbeddedSheetFilePath('UnderutilizedSubnetsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$UnderutilizedSubnetsSS = IOFactory::load($EmbeddedUnderutilizedSubnets);
					$RowNo = 2;
					foreach ($UnderutilizedSubnets->result->data as $ProviderType) {
						$UnderutilizedSubnetsS = $UnderutilizedSubnetsSS->getActiveSheet();
						$UnderutilizedSubnetsS->setCellValue('A'.$RowNo, $ProviderType->{'NetworkInsightsSubnet.address'});
						$UnderutilizedSubnetsS->setCellValue('B'.$RowNo, $ProviderType->{'NetworkInsightsSubnet.name'});
						$UnderutilizedSubnetsS->setCellValue('C'.$RowNo, $ProviderType->{'NetworkInsightsSubnet.provider'});
						$UnderutilizedSubnetsS->setCellValue('D'.$RowNo, round($ProviderType->{'NetworkInsightsSubnet.utilization_percent'},1) . '%');
						$RowNo++;
					}
					$UnderutilizedSubnetsW = IOFactory::createWriter($UnderutilizedSubnetsSS, 'Xlsx');
					$UnderutilizedSubnetsW->save($EmbeddedUnderutilizedSubnets);
				}

				// High-Risk DNS Records by Category - Slide 16
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building High-Risk DNS Records by Category");
				$EmbeddedHighRiskDNSRecordsByCategory = getEmbeddedSheetFilePath('HighRiskDNSRecordsByCategory', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$HighRiskDNSRecordsByCategorySS = IOFactory::load($EmbeddedHighRiskDNSRecordsByCategory);
				$HighRiskDNSRecordsByCategoryS = $HighRiskDNSRecordsByCategorySS->getActiveSheet();
				$RowNo = 2;
				if ($DanglingDNSCount > 0) {
					$HighRiskDNSRecordsByCategoryS->setCellValue('A'.$RowNo, 'Dangling');
					$HighRiskDNSRecordsByCategoryS->setCellValue('B'.$RowNo, $DanglingDNSCount);
					$RowNo++;
				}
				if ($AbandonedDNSCount > 0) {
					$HighRiskDNSRecordsByCategoryS->setCellValue('A'.$RowNo, 'Abandoned');
					$HighRiskDNSRecordsByCategoryS->setCellValue('B'.$RowNo, $AbandonedDNSCount);
					$RowNo++;
				}
				if ($UntrustedDNSCount > 0) {
					$HighRiskDNSRecordsByCategoryS->setCellValue('A'.$RowNo, 'Untrusted');
					$HighRiskDNSRecordsByCategoryS->setCellValue('B'.$RowNo, $UntrustedDNSCount);
					$RowNo++;
				}
				$HighRiskDNSRecordsByCategoryW = IOFactory::createWriter($HighRiskDNSRecordsByCategorySS, 'Xlsx');
				$HighRiskDNSRecordsByCategoryW->save($EmbeddedHighRiskDNSRecordsByCategory);

				// Dangling Records - Slide 16
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Dangling Records");
				$DanglingRecords = $CubeJSResults['DanglingRecords']['Body'];
				if (isset($DanglingRecords->result->data)) {
					$EmbeddedDanglingRecords = getEmbeddedSheetFilePath('DanglingRecordsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$DanglingRecordsSS = IOFactory::load($EmbeddedDanglingRecords);
					$RowNo = 2;
					foreach ($DanglingRecords->result->data as $DanglingRecord) {
						$DanglingRecordsS = $DanglingRecordsSS->getActiveSheet();
						$DanglingRecordsS->setCellValue('A'.$RowNo, $DanglingRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name'});
						$DanglingRecordsS->setCellValue('B'.$RowNo, $DanglingRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name'});
						$DanglingRecordsS->setCellValue('C'.$RowNo, $DanglingRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type'});
						$DanglingRecordsS->setCellValue('D'.$RowNo, implode(', ',$DanglingRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids'}));
						$DanglingRecordsS->setCellValue('E'.$RowNo, $DanglingRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str'});
						$RowNo++;
					}
					$DanglingRecordsW = IOFactory::createWriter($DanglingRecordsSS, 'Xlsx');
					$DanglingRecordsW->save($EmbeddedDanglingRecords);
				}

				// Abandoned Records - Slide 16
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Abandoned Records");
				$AbandonedRecords = $CubeJSResults['AbandonedRecords']['Body'];
				if (isset($AbandonedRecords->result->data)) {
					$EmbeddedAbandonedRecords = getEmbeddedSheetFilePath('AbandonedRecordsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$AbandonedRecordsSS = IOFactory::load($EmbeddedAbandonedRecords);
					$RowNo = 2;
					foreach ($AbandonedRecords->result->data as $AbandonedRecord) {
						$AbandonedRecordsS = $AbandonedRecordsSS->getActiveSheet();
						$AbandonedRecordsS->setCellValue('A'.$RowNo, $AbandonedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name'});
						$AbandonedRecordsS->setCellValue('B'.$RowNo, $AbandonedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name'});
						$AbandonedRecordsS->setCellValue('C'.$RowNo, $AbandonedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type'});
						$AbandonedRecordsS->setCellValue('D'.$RowNo, implode(', ',$AbandonedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids'}));
						$AbandonedRecordsS->setCellValue('E'.$RowNo, $AbandonedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str'});
						$RowNo++;
					}
					$AbandonedRecordsW = IOFactory::createWriter($AbandonedRecordsSS, 'Xlsx');
					$AbandonedRecordsW->save($EmbeddedAbandonedRecords);
				}

				// Untrusted Records - Slide 16
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Untrusted Records");
				$UntrustedRecords = $CubeJSResults['UntrustedRecords']['Body'];
				if (isset($UntrustedRecords->result->data)) {
					$EmbeddedUntrustedRecords = getEmbeddedSheetFilePath('UntrustedRecordsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$UntrustedRecordsSS = IOFactory::load($EmbeddedUntrustedRecords);
					$RowNo = 2;
					foreach ($UntrustedRecords->result->data as $UntrustedRecord) {
						$UntrustedRecordsS = $UntrustedRecordsSS->getActiveSheet();
						$UntrustedRecordsS->setCellValue('A'.$RowNo, $UntrustedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name'});
						$UntrustedRecordsS->setCellValue('B'.$RowNo, $UntrustedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name'});
						$UntrustedRecordsS->setCellValue('C'.$RowNo, $UntrustedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type'});
						$UntrustedRecordsS->setCellValue('D'.$RowNo, implode(', ',$UntrustedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids'}));
						$UntrustedRecordsS->setCellValue('E'.$RowNo, $UntrustedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str'});
						$RowNo++;
					}
					$UntrustedRecordsW = IOFactory::createWriter($UntrustedRecordsSS, 'Xlsx');
					$UntrustedRecordsW->save($EmbeddedUntrustedRecords);
				}

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
				$mapping = replaceTag($mapping,'#NAME01',$UserInfo->result->name);
				$mapping = replaceTag($mapping,'#NAME02',$UserInfo->result->name);
				$mapping = replaceTag($mapping,'#EMAIL01',$UserInfo->result->email);
				$mapping = replaceTag($mapping,'#EMAIL02',$UserInfo->result->email);

				##// Slide 5 - Executive Summary
				$mapping = replaceTag($mapping,'#TAG01',number_abbr($HighRiskAssetsCount)); // High-Risk Assets
				$mapping = replaceTag($mapping,'#TAG02',number_abbr($CloudSubnetOverlapCount)); // Cloud Subnet Overlap Count
				$mapping = replaceTag($mapping,'#TAG03',number_abbr($CloudSubnetUtilizationAbove50Total)); // Overutilized Subnets (>=50%)
				$mapping = replaceTag($mapping,'#TAG04',number_abbr($CloudSubnetUtilizationBelow50Total)); // Underutilized Subnets (<50%)
				$mapping = replaceTag($mapping,'#TAG05',number_abbr($DNSComplexityScore)); // DNS Complexity Score
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
				
				$mapping = replaceTag($mapping,'#TAG29',number_abbr($TotalCloudIPsCount)); // Total Allocated Cloud IPs
				$mapping = replaceTag($mapping,'#TAG30',number_abbr($AzureIPsCount)); // Total Allocated Azure IPs
				$mapping = replaceTag($mapping,'#TAG31',number_abbr($AWSIPsCount)); // Total Allocated AWS IPs
				$mapping = replaceTag($mapping,'#TAG32',number_abbr($GCPIPsCount)); // Total Allocated GCP IPs

				// $mapping = replaceTag($mapping,'#TAG29',number_abbr($CloudIPsByProviderTotalUsed['Total'])); // Total Allocated Cloud IPs
				// $mapping = replaceTag($mapping,'#TAG30',number_abbr($CloudIPsByProviderTotalUsed['Azure'])); // Total Allocated Azure IPs
				// $mapping = replaceTag($mapping,'#TAG31',number_abbr($CloudIPsByProviderTotalUsed['AWS'])); // Total Allocated AWS IPs
				// $mapping = replaceTag($mapping,'#TAG32',number_abbr($CloudIPsByProviderTotalUsed['GCP'])); // Total Allocated GCP IPs

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

				##// Slide 15 - DNS Complexity
				$mapping = replaceTag($mapping,'#TAG42',number_abbr($OverlappingZonesCount)); // Overlapping DNS Zones
				$mapping = replaceTag($mapping,'#TAG43',number_abbr($DNSComplexityScore)); // DNS Complexity Score
				$mapping = replaceTag($mapping,'#TAG44',number_abbr($CloudDNSZonesByProviderTotals['Azure'])); // Azure Zones
				$mapping = replaceTag($mapping,'#TAG45',number_abbr($CloudDNSZonesByProviderTotals['AWS'])); // AWS Zones
				$mapping = replaceTag($mapping,'#TAG46',number_abbr($CloudDNSZonesByProviderTotals['GCP'])); // GCP Zones
				$mapping = replaceTag($mapping,'#TAG47',number_abbr($CloudDNSRecordsByProviderTotals['microsoft_azure'])); // Azure Records
				$mapping = replaceTag($mapping,'#TAG48',number_abbr($CloudDNSRecordsByProviderTotals['amazon_web_service'])); // AWS Records
				$mapping = replaceTag($mapping,'#TAG49',number_abbr($CloudDNSRecordsByProviderTotals['google_cloud_platform'])); // GCP Records
				$mapping = replaceTag($mapping,'#TAG50',number_abbr($CloudForwarderCounts['Microsoft Azure']['inbound'])); // Azure Inbound Endpoints
				$mapping = replaceTag($mapping,'#TAG51',number_abbr($CloudForwarderCounts['Microsoft Azure']['outbound'])); // Azure Outbound Endpoints
				$mapping = replaceTag($mapping,'#TAG52',number_abbr($CloudForwarderCounts['Amazon Web Services']['inbound'])); // AWS Inbound Endpoints
				$mapping = replaceTag($mapping,'#TAG53',number_abbr($CloudForwarderCounts['Amazon Web Services']['outbound'])); // AWS Outbound Endpoints
				$mapping = replaceTag($mapping,'#TAG54',number_abbr($CloudForwarderCounts['Google Cloud Platform']['inbound'])); // GCP Inbound Endpoints
				$mapping = replaceTag($mapping,'#TAG55',number_abbr($CloudForwarderCounts['Google Cloud Platform']['outbound'])); // GCP Outbound Endpoints

				##// Slide 16 - High-Risk DNS Records
				$mapping = replaceTag($mapping,'#TAG56',number_abbr($HighRiskDNSRecordsCount)); // High-Risk DNS Records
				$mapping = replaceTag($mapping,'#TAG57',number_abbr($DanglingDNSCount)); // High-Risk DNS Records - Dangling
				$mapping = replaceTag($mapping,'#TAG58',number_abbr($AbandonedDNSCount)); // High-Risk DNS Records - Abandoned
				$mapping = replaceTag($mapping,'#TAG59',number_abbr($UntrustedDNSCount)); // High-Risk DNS Records - Untrusted

				##// Slide 20 - Recommendations
				$mapping = replaceTag($mapping,'#TAG60',number_abbr($TokensManagementCount)); // Management Tokens
				$mapping = replaceTag($mapping,'#TAG61',number_abbr($TokensActiveIPPercentage)); // Active IP Percentage
				$mapping = replaceTag($mapping,'#TAG62',number_abbr($ActiveIPCount)); // Active IP Count
				$mapping = replaceTag($mapping,'#TAG63',number_abbr($TokensAssetPercentage)); // Asset Count
				$mapping = replaceTag($mapping,'#TAG64',number_abbr($TokensDDIPercentage)); // DDI Objects
				$mapping = replaceTag($mapping,'#TAG65',number_abbr($TokensServer)); // Server Tokens
				$mapping = replaceTag($mapping,'#TAG66',number_abbr($TokensReporting)); // Reporting Tokens

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
			'Total' => ($Total * 23) + 17,
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