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
			echo json_encode(array(
				'result' => 'Success',
				'message' => 'Started'
			));
			fastcgi_finish_request();
	
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

			// Cloud Definitions
			$CloudAssetProviders = '["AWS","Azure","GCP", "Oracle Cloud Infrastructure"]';

			$CubeJSRequests = array(
				// All Assets
				'TotalAssetsCount' => '{"segments": [],"timeDimensions": [{"dimension": "AssetDetails_ch_agg.updated_at","dateRange": ["'.$StartDimension.'","'.$EndDimension.'"],"granularity": null}],"ungrouped": false,"dimensions": [],"measures": ["AssetDetails_ch_agg.assetsCount"], "filters": []}',
				'AssetsByClassification' => '{"measures":["AssetDetails_ch.assetsCount"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails_ch.insight_classification","AssetDetails_ch.insight_classification_label","AssetDetails_ch.insight_sub_classification","AssetDetails_ch.insight_sub_classification_label"],"filters":[{"member":"AssetDetails_ch.insight_state","operator":"equals","values":["active"]},{"member":"AssetDetails_ch.insight_classification","operator":"notEquals","values":["oseol","oseos","oseossec","deviceeol","deviceeos","deviceeossec"]}],"timezone":"UTC","order":{"AssetDetails_ch.insight_classification":"desc"},"segments":["AssetDetails_ch.existing_assets","AssetDetails_ch.asset_insights"]}',
				'AssetsByType' => '{"measures":["AssetDetails_ch.assetsCount"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails_ch.category_label"],"filters":[],"timezone":"UTC","segments":[]}',
				'AssetsByProvider' => '{"ungrouped":false,"measures":["AssetDetails_ch.assetsCount"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails_ch.provider_label"],"segments":[],"filters":[]}',
				'AssetsByLocation' => '{"ungrouped":false,"measures":["AssetDetails_ch.assetsCount"],"dimensions":["AssetDetails_ch.location_label","AssetDetails_ch.region_label","AssetDetails_ch.provider_label","AssetDetails_ch.provider_location_label"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"segments":[],"filters":[{"member":"AssetDetails_ch.region_label","operator":"notEquals","values":["unknown"]}],"limit":5}',
				'AssetReconciliation' => '{"measures":["AssetDetails_ch.assetsCount"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails_ch.reconciliation"],"filters":[{"member":"AssetDetails_ch.reconciliation_filter","operator":"contains","values":["ServiceNow"]}],"timezone":"UTC","segments":["AssetDetails_ch.existing_assets","AssetDetails_ch.managed_assets"]}',
				'AssetsMissingFromCMDB' => '{"measures":["AssetDetails_ch.assetsCount"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails_ch.providers_iterator"],"filters":[{"member":"AssetDetails_ch.providers_iterator","operator":"notContains","values":["ServiceNow"]}],"timezone":"UTC","segments":["AssetDetails_ch.managed_assets","AssetDetails_ch.existing_assets"]}',
				'AssetsOrphanedInCMDB' => '{"measures":["AssetDetails_ch.assetsCount"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails_ch.providers_iterator"],"filters":[{"member":"AssetDetails_ch.providers_iterator","operator":"contains","values":["ServiceNow"]},{"member":"AssetDetails_ch.length_providers","operator":"equals","values":[1]}],"timezone":"UTC","segments":["AssetDetails_ch.managed_assets","AssetDetails_ch.existing_assets"]}',
				'Top3AssetsMissingFromCMDB' => '{"dimensions":["AssetDetails_ch.name","AssetDetails_ch.location_label","AssetDetails_ch.ip_address","AssetDetails_ch.provider_label","AssetDetails_ch.taxonomy_type_label"],"ungrouped":true,"filters":[{"member":"AssetDetails_ch.providers_iterator","operator":"notContains","values":["ServiceNow"]}],"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null,"dimension":"AssetDetails_ch.updated_at"}],"limit":3,"segments":[],"order":{"AssetDetails_ch.category_label":"desc"},"measures":[]}',

				// Cloud Assets
				'TotalCloudAssetsCount' => '{"segments": [],"timeDimensions": [{"dimension": "AssetDetails_ch_agg.updated_at","dateRange": ["'.$StartDimension.'","'.$EndDimension.'"],"granularity": null}],"ungrouped": false,"dimensions": [],"measures": ["AssetDetails_ch_agg.assetsCount"], "filters": [{"member":"AssetDetails_ch_agg.providers_label","operator":"contains","values":'.$CloudAssetProviders.'}]}',
				'CloudAssetsByClassification' => '{"measures":["AssetDetails_ch.assetsCount"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails_ch.insight_classification","AssetDetails_ch.insight_classification_label","AssetDetails_ch.insight_sub_classification","AssetDetails_ch.insight_sub_classification_label"],"filters":[{"member":"AssetDetails_ch.insight_state","operator":"equals","values":["active"]},{"member":"AssetDetails_ch.provider_label","operator":"equals","values":'.$CloudAssetProviders.'},{"member":"AssetDetails_ch.insight_classification","operator":"notEquals","values":["oseol","oseos","oseossec","deviceeol","deviceeos","deviceeossec"]}],"timezone":"UTC","order":{"AssetDetails_ch.insight_classification":"desc"},"segments":["AssetDetails_ch.existing_assets","AssetDetails_ch.asset_insights"]}',
				'CloudAssetsWithMissingRecords' => '{"measures":["AssetDetails_ch.assetsCount"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails_ch.insight_classification_label","AssetDetails_ch.insight_indicator_label"],"filters":[{"member":"AssetDetails_ch.insight_classification","operator":"equals","values":["registration"]},{"member":"AssetDetails_ch.managed","operator":"equals","values":["true"]},{"member":"AssetDetails_ch.provider_label","operator":"equals","values":'.$CloudAssetProviders.'}],"timezone":"UTC","segments":[]}',
				'Top3CloudAssetsWithMissingRecords' => '{"dimensions":["AssetDetails_ch.name","AssetDetails_ch.vendor","AssetDetails_ch.location_label","AssetDetails_ch.ip_address","AssetDetails_ch.insight_classification_label","AssetDetails_ch.insight_indicator_label","AssetDetails_ch.last_seen","AssetDetails_ch.provider","AssetDetails_ch.category_label"],"ungrouped":true,"filters":[{"member":"AssetDetails_ch.insight_classification","operator":"equals","values":["registration"]},{"member":"AssetDetails_ch.managed","operator":"equals","values":["true"]},{"member":"AssetDetails_ch.provider_label","operator":"equals","values":'.$CloudAssetProviders.'}],"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null,"dimension":"AssetDetails_ch.updated_at"}],"limit":3,"segments":[],"order":{"AssetDetails_ch.last_seen":"desc"},"measures":[]}',
				'Top3ZombieCloudAssets' => '{"dimensions":["AssetDetails_ch.name","AssetDetails_ch.vendor","AssetDetails_ch.location_label","AssetDetails_ch.ip_address","AssetDetails_ch.insight_classification_label","AssetDetails_ch.insight_indicator_label","AssetDetails_ch.last_seen","AssetDetails_ch.provider","AssetDetails_ch.category_label"],"ungrouped":true,"filters":[{"member":"AssetDetails_ch.insight_classification","operator":"equals","values":["zombie"]},{"member":"AssetDetails_ch.managed","operator":"equals","values":["true"]},{"member":"AssetDetails_ch.provider_label","operator":"equals","values":'.$CloudAssetProviders.'}],"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null,"dimension":"AssetDetails_ch.updated_at"}],"limit":3,"segments":[],"order":{"AssetDetails_ch.updated_at":"desc"},"measures":[]}',
				'CloudAssetsByProvider' => '{"ungrouped":false,"measures":["AssetDetails_ch.assetsCount"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails_ch.provider_label"],"segments":[],"filters":[{"member":"AssetDetails_ch.provider_label","operator":"equals","values":'.$CloudAssetProviders.'}]}',
				'CloudAssetsByLocation' => '{"ungrouped":false,"measures":["AssetDetails_ch.assetsCount"],"dimensions":["AssetDetails_ch.location_label","AssetDetails_ch.region_label","AssetDetails_ch.provider_label","AssetDetails_ch.provider_location_label"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"segments":[],"filters":[{"member":"AssetDetails_ch.provider_label","operator":"equals","values":'.$CloudAssetProviders.'},{"member":"AssetDetails_ch.region_label","operator":"notEquals","values":["unknown"]}],"limit":5}',
				'CloudAssetsByType' => '{"measures":["AssetDetails_ch.assetsCount"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails_ch.category_label"],"filters":[{"member":"AssetDetails_ch.provider_label","operator":"equals","values":'.$CloudAssetProviders.'}],"timezone":"UTC","segments":[]}',
				'CloudIPsByProvider' => '{"dimensions":["AssetDetails_ch.provider_label"],"ungrouped":false,"filters":[{"member":"AssetDetails_ch.ip_address","operator":"set"},{"member":"AssetDetails_ch.provider_label","operator":"equals","values":'.$CloudAssetProviders.'}],"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null,"dimension":"AssetDetails_ch.updated_at"}],"segments":[],"measures":["AssetDetails_ch.assetsCount"]}',
				// Zombie Assets By Type - Slide 8 (Chart)	
				'CloudZombieAssetsByIndicator' => '{"ungrouped":false,"measures":["AssetDetails_ch.assetsCount"],"timeDimensions":[{"dimension":"AssetDetails_ch.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails_ch.insight_indicator_label"],"segments":[],"filters":[{"member":"AssetDetails_ch.provider_label","operator":"equals","values":'.$CloudAssetProviders.'},{"and":[{"member":"AssetDetails_ch.insight_classification","operator":"equals","values":["zombie"]},{"member":"AssetDetails_ch.managed","operator":"equals","values":["true"]}]}]}',

				// DNS Records
				'HighRiskDNSRecords' => '{"measures":["NetworkInsightsDnsRecords.count"],"timeDimensions":[{"dimension":"NetworkInsightsDnsRecords.evaluation_time","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["NetworkInsightsDnsRecords.indicator_id"],"filters":[{"member":"NetworkInsightsDnsRecords.indicator_id","operator":"set"}],"timezone":"UTC","segments":[]}',

				// Licensing
				'LicensingManagement' => '{"measures":["TokenUtilManagementObjects.count"],"timeDimensions":[{"dimension":"TokenUtilManagementObjects.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":"day"}],"dimensions":["TokenUtilManagementObjects.category","TokenUtilManagementObjects.object_type"],"filters":[{"and":[{"member":"TokenUtilManagementObjects.object_type","operator":"equals","values":["DDI","Active IPs","Assets"]},{"member":"TokenUtilManagementObjects.category","operator":"equals","values":["Native"]}]}]}',
				'LicensingServer' => '{"measures":["TokenUtilProtoSrvSM.count","TokenUtilProtoSrvSM.tokens"],"timeDimensions":[{"dimension":"TokenUtilProtoSrvSM.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":"day"}],"dimensions":[]}',
				'LicensingReporting' => '{"measures":["TokenUtilReporting.count","TokenUtilReporting.tokens"],"timeDimensions":[{"dimension":"TokenUtilReporting.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":"day"}],"dimensions":[]}',

				// Original
				// Verified - Some updates may be required when OCI gets added to NetworkInsightsSubnet, NetworkInsightsDnsRecords & NetworkInsightsDnsRecordsRolledByIndicatorId cubes
				'CloudSubnetUtilizationAbove50' => '{"measures":[],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]},{"operator":"gte","member":"NetworkInsightsSubnet.utilization_percent","values":["50"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.ref_id_resource_id","NetworkInsightsSubnet.address","NetworkInsightsSubnet.utilization_percent","NetworkInsightsSubnet.provider","NetworkInsightsSubnet.usage","NetworkInsightsSubnet.name"],"timeDimensions":[{"dimension":"NetworkInsightsSubnet.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"limit":3}',
				'CloudSubnetUtilizationAbove50Count' => '{"measures":["NetworkInsightsSubnet.count"],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]},{"operator":"gte","member":"NetworkInsightsSubnet.utilization_percent","values":["50"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.provider"],"timeDimensions":[{"dimension":"NetworkInsightsSubnet.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}]}',
				'CloudSubnetUtilizationBelow50' => '{"measures":[],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]},{"operator":"lt","member":"NetworkInsightsSubnet.utilization_percent","values":["50"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.ref_id_resource_id","NetworkInsightsSubnet.address","NetworkInsightsSubnet.utilization_percent","NetworkInsightsSubnet.provider","NetworkInsightsSubnet.usage","NetworkInsightsSubnet.name"],"timeDimensions":[{"dimension":"NetworkInsightsSubnet.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"limit":3}',
				'CloudSubnetUtilizationBelow50Count' => '{"measures":["NetworkInsightsSubnet.count"],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]},{"operator":"lt","member":"NetworkInsightsSubnet.utilization_percent","values":["50"]}],"ungrouped":false,"dimensions":["NetworkInsightsSubnet.provider"],"timeDimensions":[{"dimension":"NetworkInsightsSubnet.updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}]}',
				'CloudSubnetsByProvider' => '{"dimensions":["NetworkInsightsSubnet.provider"],"ungrouped":false,"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"dimension":"NetworkInsightsSubnet.updated_at","granularity":null}],"measures":["NetworkInsightsSubnet.count"],"segments":[],"filters":[{"operator":"equals","member":"NetworkInsightsSubnet.source","values":["Cloud"]}]}',
				'CloudDNSRecordsByProvider' => '{"measures":["NetworkInsightsDnsRecords.count"],"timeDimensions":[{"dimension":"NetworkInsightsDnsRecords.evaluation_time","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"segments":[],"ungrouped":false,"dimensions":["NetworkInsightsDnsRecords.provider"],"filters":[{"member":"NetworkInsightsDnsRecords.provider","operator":"equals","values":["amazon_web_service","microsoft_azure","google_cloud_platform","oracle_cloud_infrastructure"]}],"timezone":"UTC"}',
				'Top3AbandonedRecords' => '{"dimensions":["NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type","NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str"],"filters":[{"member":"NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids_str","operator":"contains","values":["Abandoned"]}],"timeDimensions":[{"dimension":"NetworkInsightsDnsRecordsRolledByIndicatorId.evaluation_time","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null}],"limit":3,"offset":0,"order":{"NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str":"asc"},"ungrouped":true,"total":true}',
				'Top3UntrustedRecords' => '{"dimensions":["NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type","NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str"],"filters":[{"member":"NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids_str","operator":"contains","values":["Untrusted"]}],"timeDimensions":[{"dimension":"NetworkInsightsDnsRecordsRolledByIndicatorId.evaluation_time","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null}],"limit":3,"offset":0,"order":{"NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str":"asc"},"ungrouped":true,"total":true}',
				'Top3DanglingRecords' => '{"dimensions":["NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type","NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids","NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str"],"filters":[{"member":"NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids_str","operator":"contains","values":["Dangling"]}],"timeDimensions":[{"dimension":"NetworkInsightsDnsRecordsRolledByIndicatorId.evaluation_time","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null}],"limit":3,"offset":0,"order":{"NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str":"asc"},"ungrouped":true,"total":true}',
				'OverlappingSubnets' => '{"dimensions":["NetworkInsightsOverlappingBlocksRolled.address","NetworkInsightsOverlappingBlocksRolled.overlap_count","NetworkInsightsOverlappingBlocksRolled.sources","NetworkInsightsOverlappingBlocksRolled.providers"],"filters":[{"member":"NetworkInsightsOverlappingBlocksRolled.sources_str","operator":"contains","values":["Cloud"]}],"timeDimensions":[],"limit":3,"offset":0,"total":true,"order":{"NetworkInsightsOverlappingBlocksRolled.overlap_count":"desc"}}',

				// Not returned via new Cube AssetDetails_ch, and unsure of the field name for Oracle Cloud Infrastructure
				'CloudDNSZonesByProvider' => '{"ungrouped":false,"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails.provider_label","AssetDetails.doc_asset_category"],"segments":[],"filters":[{"and":[{"member":"AssetDetails.doc_asset_category","operator":"equals","values":["dns"]},{"member":"AssetDetails.provider_label","operator":"equals","values":'.$CloudAssetProviders.'},{"member":"AssetDetails.doc_asset_display_type","operator":"equals","values":["aws_route53_hosted_zones","azure_dns_zones","azure_privatedns_private_zones","gcp_dns_managed_zones"]}]}]}',
				'OverlappingZones' => '{"ungrouped":false,"measures":["AssetDetails.count"],"timeDimensions":[{"dimension":"AssetDetails.doc_updated_at","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["AssetDetails.provider_label","AssetDetails.doc_asset_category","AssetDetails.doc_name"],"segments":[],"filters":[{"and":[{"member":"AssetDetails.doc_asset_category","operator":"equals","values":["dns"]},{"member":"AssetDetails.provider_label","operator":"equals","values":'.$CloudAssetProviders.'},,{"member":"AssetDetails.resource_type","operator":"equals","values":["aws_route53_hosted_zones","azure_dns_zones","azure_privatedns_private_zones","gcp_dns_managed_zones"]}]}]}'
			);

			$CubeJSResults = $this->QueryCubeJSMulti($CubeJSRequests);
	
			// Extract Powerpoint Template(s) as Zip
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Extracting template(s)");
			foreach ($SelectedTemplates as &$SelectedTemplate) {
				$macroEnabled = $SelectedTemplate['macroEnabled'] ?? false;
				if ($macroEnabled) {
					$SelectedTemplateFileExt = '.pptm';
				} else {
					$SelectedTemplateFileExt = '.pptx';
				}

				$ExtractedDir = $this->getDir()['Files'].'/reports/'.str_replace($SelectedTemplateFileExt,'', 'report'.'-'.$config['UUID'].'-'.$SelectedTemplate['FileName']);

				error_log("Extracting template: ".$SelectedTemplate['FileName']." to: ".$ExtractedDir);

				extractZip($this->getDir()['Files'].'/templates/'.$SelectedTemplate['FileName'],$ExtractedDir);
				$SelectedTemplate['ExtractedDir'] = $ExtractedDir;
			}


			// Define the embedded sheets with their corresponding file numbers
			// This needs to match across all active templates at this moment
			$EmbeddedSheets = [
				'CloudAssetsByProvider' => 0, // Slide 7
				'CloudAssetsByLocation' => 1, // Slide 7
				'CloudZombieAssetsByIndicator' => array(
					'Landscape' => 3,
					'Portrait' => 2
				), // Slide 8
				'CloudAssetsByType' => array(
					'Landscape' => 2,
					'Portrait' => 3
				), // Slide 8
				'ZombieAssetsTable' => 4, // Slide 8
				'CloudAssetsWithMissingRecords' => 5, // Slide 9
				'CloudAssetsNonCompliant' => 6, // Slide 9
				'CloudAssetsWithMissingRecordsTable' => 7, // Slide 9
				// New Slides
				'AssetsByProvider' => 8, // Slide 11
				'AssetsByLocation' => 9, // Slide 11
				'AssetsByType' => 10, // Slide 11
				'AssetsMissingFromCMDB' => 11, // Slide 13
				'AssetsOrphanedInCMDB' => 12, // Slide 13
				'AssetsMissingFromCMDBTable' => 13, // Slide 13
				// Slide 14 - 14, 15, 16 - TODO
				// New Slides
				'TotalIPsByProvider' => 17, // Slide 11
				'OverlappingSubnetsTable' => 18, // Slide 11
				'OverutilizedSubnetsTable' => 19, // Slide 13
				'UnderutilizedSubnetsTable' => 20, // Slide 13
				'HighRiskDNSRecordsByCategory' => 21, // Slide 16
				'Top3DanglingRecordsTable' => 22, // Slide 16
				'Top3AbandonedRecordsTable' => 23, // Slide 17
				'Top3UntrustedRecordsTable' => 24 // Slide 17
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

			// Total Assets
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Total Assets");
			$TotalCloudAssetsCount = $CubeJSResults['TotalCloudAssetsCount']['Body']->result->data[0]->{'AssetDetails_ch_agg.assetsCount'} ?? 0;
			$TotalAssetsCount = $CubeJSResults['TotalAssetsCount']['Body']->result->data[0]->{'AssetDetails_ch_agg.assetsCount'} ?? 0;

			// All Assets By Classification
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building All Assets by Classification");
			$AssetsByClassification = $CubeJSResults['AssetsByClassification']['Body']->result->data ?? array();
			$GhostAssetsCount = 0;
			$GhostAssets = array();
			$ZombieAssetsCount = 0;
			$ZombieAssets = array();
			$NonCompliantAssetsCount = 0;
			$NonCompliantAssets = array();
			$UnregisteredAssetsCount = 0;
			$UnregisteredAssets = array();
			$HighRiskAssetsCount = 0;
			foreach ($AssetsByClassification as $value) {
				$HighRiskAssetsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
				switch ($value->{'AssetDetails_ch.insight_classification'}) {
					case 'ghost':
						$GhostAssetsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						$GhostAssets[$value->{'AssetDetails_ch.insight_sub_classification'}] = $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						break;
					case 'zombie':
						$ZombieAssetsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						$ZombieAssets[$value->{'AssetDetails_ch.insight_sub_classification'}] = $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						break;
					case 'compliance':
						$NonCompliantAssetsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						$NonCompliantAssets[$value->{'AssetDetails_ch.insight_sub_classification'}] = $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						break;
					case 'registration':
						$UnregisteredAssetsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						$UnregisteredAssets[$value->{'AssetDetails_ch.insight_sub_classification'}] = $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						break;
				}
			}
			if ($ZombieAssetsCount != 0) {
				$ZombieAssetsResourceUtilizationIdlePerc = ($ZombieAssets['resourceutilizationidle'] ?? 0 / $ZombieAssetsCount) * 100;
				$ZombieAssetsResourceUtilizationLowPerc = ($ZombieAssets['resourceutilizationlow'] ?? 0 / $ZombieAssetsCount) * 100;
				$ZombieAssetsOrphanedPerc = ($ZombieAssets['orphan'] ?? 0 / $ZombieAssetsCount) * 100;
			}

			// Cloud Assets By Classification
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Cloud Assets by Classification");
			$GhostCloudAssetsCount = 0;
			$ZombieCloudAssetsCount = 0;
			$ZombieCloudAssets = array();
			$NonCompliantCloudAssetsCount = 0;
			$NonCompliantCloudAssets = array();
			$UnregisteredCloudAssetsCount = 0;
			$UnregisteredCloudAssets = array();
			$HighRiskCloudAssetsCount = 0;
			$CloudAssetsByClassification = $CubeJSResults['CloudAssetsByClassification']['Body']->result->data ?? array();
			foreach ($CloudAssetsByClassification as $value) {
				$HighRiskCloudAssetsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
				switch ($value->{'AssetDetails_ch.insight_classification'}) {
					case 'ghost':
						$GhostCloudAssetsCount += $value->{'AssetDetails_ch.assetsCount'};
						break;
					case 'zombie':
						$ZombieCloudAssetsCount += $value->{'AssetDetails_ch.assetsCount'};
						$ZombieCloudAssets[$value->{'AssetDetails_ch.insight_sub_classification'}] = $value->{'AssetDetails_ch.assetsCount'};
						break;
					case 'compliance':
						$NonCompliantCloudAssetsCount += $value->{'AssetDetails_ch.assetsCount'};
						$NonCompliantCloudAssets[$value->{'AssetDetails_ch.insight_sub_classification_label'}] = $value->{'AssetDetails_ch.assetsCount'};
						break;
					case 'registration':
						$UnregisteredCloudAssetsCount += $value->{'AssetDetails_ch.assetsCount'};
						$UnregisteredCloudAssets[$value->{'AssetDetails_ch.insight_sub_classification'}] = $value->{'AssetDetails_ch.assetsCount'};
						break;
				}
			}
			if ($ZombieCloudAssetsCount != 0) {
				$ZombieCloudAssetsResourceUtilizationIdlePerc = ($ZombieCloudAssets['resourceutilizationidle'] ?? 0 / $ZombieCloudAssetsCount) * 100;
				$ZombieCloudAssetsResourceUtilizationLowPerc = ($ZombieCloudAssets['resourceutilizationlow'] ?? 0 / $ZombieCloudAssetsCount) * 100;
				$ZombieCloudAssetsOrphanedPerc = ($ZombieCloudAssets['orphan'] ?? 0 / $ZombieCloudAssetsCount) * 100;
			}

			// Subnet Utilization
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Subnet Utilization");
			$CloudSubnetUtilizationAbove50Count = $CubeJSResults['CloudSubnetUtilizationAbove50Count']['Body']->result->data ?? array();
			$CloudSubnetUtilizationAbove50ByProvider = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Oracle Cloud Infrastructure" => 0, "Total" => 0];
			$CloudSubnetUtilizationAbove50Total = 0;
			foreach ($CloudSubnetUtilizationAbove50Count as $value) {
				$CloudSubnetUtilizationAbove50ByProvider[$value->{'NetworkInsightsSubnet.provider'}] = $value->{'NetworkInsightsSubnet.count'} ?? 0;
				$CloudSubnetUtilizationAbove50Total += $value->{'NetworkInsightsSubnet.count'} ?? 0;
			}
			$CloudSubnetUtilizationBelow50Count = $CubeJSResults['CloudSubnetUtilizationBelow50Count']['Body']->result->data ?? array();
			$CloudSubnetUtilizationBelow50ByProvider = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Oracle Cloud Infrastructure" => 0, "Total" => 0];
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
			$OCISubnetsCount = 0;
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
					case 'Oracle Cloud Infrastructure':
						$OCISubnetsCount += $value->{'NetworkInsightsSubnet.count'} ?? 0;
						$TotalCloudSubnetsCount += $value->{'NetworkInsightsSubnet.count'} ?? 0;
						break;
				}
				$TotalSubnetsCount += $value->{'NetworkInsightsSubnet.count'} ?? 0;
			}
			if ($TotalCloudSubnetsCount != 0) {
				$AWSSubnetsPercentage = ($AWSSubnetsCount / $TotalCloudSubnetsCount) * 100;
				$AzureSubnetsPercentage = ($AzureSubnetsCount / $TotalCloudSubnetsCount) * 100;
				$GCPSubnetsPercentage = ($GCPSubnetsCount / $TotalCloudSubnetsCount) * 100;
				$OCISubnetsPercentage = ($OCISubnetsCount / $TotalCloudSubnetsCount) * 100;
			} else {
				$AWSSubnetsPercentage = 0;
				$AzureSubnetsPercentage = 0;
				$GCPSubnetsPercentage = 0;
				$OCISubnetsPercentage = 0;
			}

			// IPs by Provider
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building IPs by Provider");
			$CloudIPsByProvider = $CubeJSResults['CloudIPsByProvider']['Body']->result->data ?? array();

			$AWSIPsCount = 0;
			$AzureIPsCount = 0;
			$GCPIPsCount = 0;
			$OCIIPsCount = 0;
			$TotalCloudIPsCount = 0;
			foreach ($CloudIPsByProvider as $value) {
				switch ($value->{'AssetDetails_ch.provider_label'}) {
					case 'AWS':
						$AWSIPsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						$TotalCloudIPsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						break;
					case 'Azure':
						$AzureIPsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						$TotalCloudIPsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						break;
					case 'GCP':
						$GCPIPsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						$TotalCloudIPsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						break;
					case 'Oracle Cloud Infrastructure':
						$OCIIPsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						$TotalCloudIPsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
						break;
				}
				$TotalSubnetsCount += $value->{'AssetDetails_ch.assetsCount'} ?? 0;
			}

			// DNS Zones by Provider
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building DNS Zones by Provider");
			$CloudDNSZonesByProvider = $CubeJSResults['CloudDNSZonesByProvider']['Body']->result->data ?? array();
			$CloudDNSZonesByProviderTotals = ["AWS" => 0, "Azure" => 0, "GCP" => 0, "Oracle Cloud Infrastructure" => 0, "Total" => 0];
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
			$CloudDNSRecordsByProviderTotals = ["amazon_web_service" => 0, "microsoft_azure" => 0, "google_cloud_platform" => 0, "oracle_cloud_infrastructure" => 0, "Total" => 0];
			
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

			// Cloud Forwarders
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Cloud Forwarders");
			$CloudForwardersResult = $this->QueryCSP("get","api/dns/v1/cloud_forwarder/endpoint");
			$CloudForwarders = $CloudForwardersResult->results ?? array();
			$CloudForwarderCounts = [
				'Microsoft Azure' => ['inbound' => 0, 'outbound' => 0],
				'Amazon Web Services' => ['inbound' => 0, 'outbound' => 0],
				'Google Cloud Platform' => ['inbound' => 0, 'outbound' => 0],
				'Oracle Cloud Infrastructure' => ['inbound' => 0, 'outbound' => 0],
				'Total' => 0
			];
				
			foreach ($CloudForwarders as $CloudForwarder) {
				$provider = $CloudForwarder->provider_type;
				$direction = $CloudForwarder->direction ?? 'unknown';
				if ($direction != 'unknown') {
					$CloudForwarderCounts[$provider][$direction]++;
				}
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
			$providers = ['AWS', 'Azure', 'GCP', 'Oracle Cloud Infrastructure'];
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
			foreach (['amazon_web_service', 'microsoft_azure', 'google_cloud_platform', 'oracle_cloud_infrastructure'] as $provider) {
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

			// CMDB - Service Now Reconciliation
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building CMDB Reconciliation");
			$AssetReconciliation = $CubeJSResults['AssetReconciliation']['Body']->result->data ?? array();
			$AssetsOnlyInServiceNow = 0;
			$AssetsMissingFromServiceNow = 0;
			$AssetsMatchingInServiceNow = 0;
			
			foreach ($AssetReconciliation as $CMDBAsset) {
				switch ($CMDBAsset->{'AssetDetails_ch.reconciliation'}) {
					case 'Only in ServiceNow':
						$AssetsOnlyInServiceNow += $CMDBAsset->{'AssetDetails_ch.assetsCount'};
						break;
					case 'Missing From ServiceNow':
						$AssetsMissingFromServiceNow += $CMDBAsset->{'AssetDetails_ch.assetsCount'};
						break;
					case 'Matching':
						$AssetsMatchingInServiceNow += $CMDBAsset->{'AssetDetails_ch.assetsCount'};
						break;
				}
			}

			$AssetsOnlyInServiceNowPerc = $AssetsOnlyInServiceNow > 0 ? ($AssetsOnlyInServiceNow / $TotalAssetsCount) * 100 : 0;
			$AssetsMissingFromServiceNowPerc = $AssetsMissingFromServiceNow > 0 ? ($AssetsMissingFromServiceNow / $TotalAssetsCount) * 100 : 0;
			$AssetsMatchingInServiceNowPerc = $AssetsMatchingInServiceNow > 0 ? ($AssetsMatchingInServiceNow / $TotalAssetsCount) * 100 : 0;

			// Loop for each selected template
			foreach ($SelectedTemplates as &$SelectedTemplate) {
				$embeddedDirectory = $SelectedTemplate['ExtractedDir'].'/ppt/embeddings/';
				$embeddedFiles = array_values(array_diff(scandir($embeddedDirectory), array('.', '..')));
				usort($embeddedFiles, 'strnatcmp');

				// Is Macro Enabled?
				$macroEnabled = $SelectedTemplate['macroEnabled'] ?? false;
				if ($macroEnabled) {
					$SelectedTemplateFileExt = '.pptm';
				} else {
					$SelectedTemplateFileExt = '.pptx';
				}

				$this->logging->writeLog("Assessment","Embedded Files List","debug",['Template' => $SelectedTemplate, 'Embedded Files' => $embeddedFiles]);


				// ** CHARTS & TABLES ** //
				// Cloud Assets By Provider - Slide 7
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Cloud Assets by Provider");
				$CloudAssetsByProvider = $CubeJSResults['CloudAssetsByProvider']['Body'];
				if (isset($CloudAssetsByProvider->result->data)) {
					$EmbeddedCloudAssetsByProvider = getEmbeddedSheetFilePath('CloudAssetsByProvider', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$CloudAssetsByProviderSS = IOFactory::load($EmbeddedCloudAssetsByProvider);
					$RowNo = 2;
					foreach ($CloudAssetsByProvider->result->data as $ProviderType) {
						$CloudAssetsByProviderS = $CloudAssetsByProviderSS->getActiveSheet();
						$CloudAssetsByProviderS->setCellValue('A'.$RowNo, $this->convertProvider($ProviderType->{'AssetDetails_ch.provider_label'}));
						$CloudAssetsByProviderS->setCellValue('B'.$RowNo, $ProviderType->{'AssetDetails_ch.assetsCount'});
						$RowNo++;
					}
					$CloudAssetsByProviderW = IOFactory::createWriter($CloudAssetsByProviderSS, 'Xlsx');
					$CloudAssetsByProviderW->save($EmbeddedCloudAssetsByProvider);
				}

				// Cloud Assets By Location - Slide 7
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Cloud Assets by Location");
				$CloudAssetsByLocation = $CubeJSResults['CloudAssetsByLocation']['Body'];
				if (isset($CloudAssetsByLocation->result->data)) {
					$EmbeddedCloudAssetsByLocation = getEmbeddedSheetFilePath('CloudAssetsByLocation', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$CloudAssetsByLocationSS = IOFactory::load($EmbeddedCloudAssetsByLocation);
					$RowNo = 2;
					$CloudAssetsByLocationIncludedCount = 0;
					foreach ($CloudAssetsByLocation->result->data as $AssetLocation) {
						$CloudAssetsByLocationS = $CloudAssetsByLocationSS->getActiveSheet();
						$CloudAssetsByLocationS->setCellValue('A'.$RowNo, $this->convertProvider($AssetLocation->{'AssetDetails_ch.provider_label'}).': '.$AssetLocation->{'AssetDetails_ch.location_label'});
						$CloudAssetsByLocationS->setCellValue('B'.$RowNo, round((100 / $TotalCloudAssetsCount) * $AssetLocation->{'AssetDetails_ch.assetsCount'},2) / 100);
						$CloudAssetsByLocationIncludedCount += $AssetLocation->{'AssetDetails_ch.assetsCount'};
						$RowNo++;
					}
					$CloudAssetsByLocationExcludedCount = $TotalCloudAssetsCount - $CloudAssetsByLocationIncludedCount;
					if ($CloudAssetsByLocationExcludedCount > 0) {
						$CloudAssetsByLocationS = $CloudAssetsByLocationSS->getActiveSheet();
						$CloudAssetsByLocationS->setCellValue('A'.$RowNo, 'Others');
						$CloudAssetsByLocationS->setCellValue('B'.$RowNo, round((100 / $TotalCloudAssetsCount) * $CloudAssetsByLocationExcludedCount,2) / 100);
					}
					$CloudAssetsByLocationW = IOFactory::createWriter($CloudAssetsByLocationSS, 'Xlsx');
					$CloudAssetsByLocationW->save($EmbeddedCloudAssetsByLocation);
				}

				// Cloud Assets By Type - Slide 8
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Cloud Assets by Type");
				$CloudAssetsByType = $CubeJSResults['CloudAssetsByType']['Body'];
				if (isset($CloudAssetsByType->result->data)) {
					$EmbeddedCloudAssetsByType = getEmbeddedSheetFilePath('CloudAssetsByType', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$CloudAssetsByTypeSS = IOFactory::load($EmbeddedCloudAssetsByType);
					$RowNo = 2;
					foreach ($CloudAssetsByType->result->data as $AssetCategory) {
						$CloudAssetsByTypeS = $CloudAssetsByTypeSS->getActiveSheet();
						$CloudAssetsByTypeS->setCellValue('A'.$RowNo, $AssetCategory->{'AssetDetails_ch.category_label'});
						$CloudAssetsByTypeS->setCellValue('B'.$RowNo, $AssetCategory->{'AssetDetails_ch.assetsCount'});
						$RowNo++;
					}
					$CloudAssetsByTypeW = IOFactory::createWriter($CloudAssetsByTypeSS, 'Xlsx');
					$CloudAssetsByTypeW->save($EmbeddedCloudAssetsByType);
				}

				// Zombie Assets by Type Chart - Slide 8
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Zombie Assets by Indicator");
				$CloudZombieAssetsByIndicator = $CubeJSResults['CloudZombieAssetsByIndicator']['Body'];

				if (isset($CloudZombieAssetsByIndicator->result->data)) {
					$EmbeddedCloudZombieAssetsByIndicator = getEmbeddedSheetFilePath('CloudZombieAssetsByIndicator', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$CloudZombieAssetsByIndicatorSS = IOFactory::load($EmbeddedCloudZombieAssetsByIndicator);
					$RowNo = 2;
					foreach ($CloudZombieAssetsByIndicator->result->data as $ZombieAssetIndicator) {
						$CloudZombieAssetsByIndicatorS = $CloudZombieAssetsByIndicatorSS->getActiveSheet();
						$CloudZombieAssetsByIndicatorS->setCellValue('A'.$RowNo, $ZombieAssetIndicator->{'AssetDetails_ch.insight_indicator_label'});
						$CloudZombieAssetsByIndicatorS->setCellValue('B'.$RowNo, $ZombieAssetIndicator->{'AssetDetails_ch.assetsCount'});
						$RowNo++;
					}
					$CloudZombieAssetsByIndicatorW = IOFactory::createWriter($CloudZombieAssetsByIndicatorSS, 'Xlsx');
					$CloudZombieAssetsByIndicatorW->save($EmbeddedCloudZombieAssetsByIndicator);
				}

				// Zombie Assets Table - Slide 8
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Zombie Assets Table");
				$Top3ZombieCloudAssets = $CubeJSResults['Top3ZombieCloudAssets']['Body'];
				if (isset($Top3ZombieCloudAssets->result->data)) {
					$EmbeddedZombieAssetsTable = getEmbeddedSheetFilePath('ZombieAssetsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$ZombieAssetsTableSS = IOFactory::load($EmbeddedZombieAssetsTable);
					$RowNo = 2;
					foreach ($Top3ZombieCloudAssets->result->data as $Top3ZombieAsset) {
						$ZombieAssetsTableS = $ZombieAssetsTableSS->getActiveSheet();
						$ZombieAssetsTableS->setCellValue('A'.$RowNo, $Top3ZombieAsset->{'AssetDetails_ch.name'}); // Asset Name
						$ZombieAssetsTableS->setCellValue('B'.$RowNo, $Top3ZombieAsset->{'AssetDetails_ch.category_label'}); // Asset Category / Type
						$ZombieAssetsTableS->setCellValue('C'.$RowNo, $Top3ZombieAsset->{'AssetDetails_ch.location_label'}); // Asset Location
						$ZombieAssetsTableS->setCellValue('D'.$RowNo, $Top3ZombieAsset->{'AssetDetails_ch.ip_address'}); // Asset IP Address
						$ZombieAssetsTableS->setCellValue('E'.$RowNo, $Top3ZombieAsset->{'AssetDetails_ch.insight_classification_label'}); // Asset Classification
						$ZombieAssetsTableS->setCellValue('F'.$RowNo, $Top3ZombieAsset->{'AssetDetails_ch.insight_indicator_label'}); // Asset Indicator
						$ZombieAssetsTableS->setCellValue('G'.$RowNo, $this->timeAgo($Top3ZombieAsset->{'AssetDetails_ch.last_seen'})); // Asset Last Seen
						$ZombieAssetsTableS->setCellValue('H'.$RowNo, $Top3ZombieAsset->{'AssetDetails_ch.provider'}); // Asset Provider
						$RowNo++;
					}
					$ZombieAssetsTableW = IOFactory::createWriter($ZombieAssetsTableSS, 'Xlsx');
					$ZombieAssetsTableW->save($EmbeddedZombieAssetsTable);
				}

				// Cloud Assets with Missing Records Chart - Slide 9
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Cloud Assets with Missing Records");
				$CloudAssetsWithMissingRecords = $CubeJSResults['CloudAssetsWithMissingRecords']['Body'];
				if (isset($CloudAssetsWithMissingRecords->result->data)) {
					$EmbeddedCloudAssetsWithMissingRecords = getEmbeddedSheetFilePath('CloudAssetsWithMissingRecords', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$CloudAssetsWithMissingRecordsSS = IOFactory::load($EmbeddedCloudAssetsWithMissingRecords);
					$RowNo = 2;
					foreach ($CloudAssetsWithMissingRecords->result->data as $CloudAssetsWithMissingRecordsIndicatorType) {
						$CloudAssetsWithMissingRecordsS = $CloudAssetsWithMissingRecordsSS->getActiveSheet();
						$CloudAssetsWithMissingRecordsS->setCellValue('A'.$RowNo, $CloudAssetsWithMissingRecordsIndicatorType->{'AssetDetails_ch.insight_indicator_label'});
						$CloudAssetsWithMissingRecordsS->setCellValue('B'.$RowNo, $CloudAssetsWithMissingRecordsIndicatorType->{'AssetDetails_ch.assetsCount'});
						$RowNo++;
					}
					$CloudAssetsWithMissingRecordsW = IOFactory::createWriter($CloudAssetsWithMissingRecordsSS, 'Xlsx');
					$CloudAssetsWithMissingRecordsW->save($EmbeddedCloudAssetsWithMissingRecords);
				}

				// Cloud Assets with Missing Records Table - Slide 9
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Cloud Assets with Missing Records Table");
				$Top3CloudAssetsWithMissingRecords = $CubeJSResults['Top3CloudAssetsWithMissingRecords']['Body'];
				if (isset($Top3CloudAssetsWithMissingRecords->result->data)) {
					$EmbeddedCloudAssetsWithMissingRecordsTable = getEmbeddedSheetFilePath('CloudAssetsWithMissingRecordsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$CloudAssetsWithMissingRecordsTableSS = IOFactory::load($EmbeddedCloudAssetsWithMissingRecordsTable);
					$RowNo = 2;
					foreach ($Top3CloudAssetsWithMissingRecords->result->data as $Top3CloudAssetWithMissingRecords) {
						$CloudAssetsWithMissingRecordsTableS = $CloudAssetsWithMissingRecordsTableSS->getActiveSheet();
						$CloudAssetsWithMissingRecordsTableS->setCellValue('A'.$RowNo, $Top3CloudAssetWithMissingRecords->{'AssetDetails_ch.name'}); // Asset Name
						$CloudAssetsWithMissingRecordsTableS->setCellValue('B'.$RowNo, $Top3CloudAssetWithMissingRecords->{'AssetDetails_ch.category_label'}); // Asset Category / Type
						$CloudAssetsWithMissingRecordsTableS->setCellValue('C'.$RowNo, $Top3CloudAssetWithMissingRecords->{'AssetDetails_ch.location_label'}); // Asset Location
						$CloudAssetsWithMissingRecordsTableS->setCellValue('D'.$RowNo, $Top3CloudAssetWithMissingRecords->{'AssetDetails_ch.ip_address'} ?? null); // Asset IP Address
						$CloudAssetsWithMissingRecordsTableS->setCellValue('E'.$RowNo, $Top3CloudAssetWithMissingRecords->{'AssetDetails_ch.insight_classification_label'}); // Asset Classification
						$CloudAssetsWithMissingRecordsTableS->setCellValue('F'.$RowNo, $Top3CloudAssetWithMissingRecords->{'AssetDetails_ch.insight_indicator_label'}); // Asset Indicator
						$CloudAssetsWithMissingRecordsTableS->setCellValue('G'.$RowNo, $this->timeAgo($Top3CloudAssetWithMissingRecords->{'AssetDetails_ch.last_seen'})); // Asset Last Seen
						$CloudAssetsWithMissingRecordsTableS->setCellValue('H'.$RowNo, $Top3CloudAssetWithMissingRecords->{'AssetDetails_ch.provider'}); // Asset Provider
						$RowNo++;
					}
					$CloudAssetsWithMissingRecordsTableW = IOFactory::createWriter($CloudAssetsWithMissingRecordsTableSS, 'Xlsx');
					$CloudAssetsWithMissingRecordsTableW->save($EmbeddedCloudAssetsWithMissingRecordsTable);
				}


				// Non-Compliant Cloud Assets Chart - Slide 9
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Non-Compliant Cloud Assets");
				if (!empty($NonCompliantCloudAssets)) {
					$EmbeddedNonCompliantCloudAssets = getEmbeddedSheetFilePath('CloudAssetsNonCompliant', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$NonCompliantCloudAssetsSS = IOFactory::load($EmbeddedNonCompliantCloudAssets);
					$RowNo = 2;
					foreach ($NonCompliantCloudAssets as $InsightType => $Count) {
						$NonCompliantCloudAssetsS = $NonCompliantCloudAssetsSS->getActiveSheet();
						$NonCompliantCloudAssetsS->setCellValue('A'.$RowNo, $InsightType);
						$NonCompliantCloudAssetsS->setCellValue('B'.$RowNo, $Count);
						$RowNo++;
					}
					$NonCompliantCloudAssetsW = IOFactory::createWriter($NonCompliantCloudAssetsSS, 'Xlsx');
					$NonCompliantCloudAssetsW->save($EmbeddedNonCompliantCloudAssets);
				}

				// New Slides
				// All Assets By Type - Slide 11
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Assets by Type");
				$AssetsByType = $CubeJSResults['AssetsByType']['Body'];
				if (isset($AssetsByType->result->data)) {
					$EmbeddedAssetsByType = getEmbeddedSheetFilePath('AssetsByType', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$AssetsByTypeSS = IOFactory::load($EmbeddedAssetsByType);
					$RowNo = 2;
					foreach ($AssetsByType->result->data as $AssetCategory) {
						$AssetsByTypeS = $AssetsByTypeSS->getActiveSheet();
						$AssetsByTypeS->setCellValue('A'.$RowNo, $AssetCategory->{'AssetDetails_ch.category_label'});
						$AssetsByTypeS->setCellValue('B'.$RowNo, $AssetCategory->{'AssetDetails_ch.assetsCount'});
						$RowNo++;
					}
					$AssetsByTypeW = IOFactory::createWriter($AssetsByTypeSS, 'Xlsx');
					$AssetsByTypeW->save($EmbeddedAssetsByType);
				}

				// Assets By Provider - Slide 11
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Assets by Provider");
				$AssetsByProvider = $CubeJSResults['AssetsByProvider']['Body'];
				if (isset($AssetsByProvider->result->data)) {
					$EmbeddedAssetsByProvider = getEmbeddedSheetFilePath('AssetsByProvider', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$AssetsByProviderSS = IOFactory::load($EmbeddedAssetsByProvider);
					$RowNo = 2;
					foreach ($AssetsByProvider->result->data as $ProviderType) {
						$AssetsByProviderS = $AssetsByProviderSS->getActiveSheet();
						$AssetsByProviderS->setCellValue('A'.$RowNo, $this->convertProvider($ProviderType->{'AssetDetails_ch.provider_label'}));
						$AssetsByProviderS->setCellValue('B'.$RowNo, $ProviderType->{'AssetDetails_ch.assetsCount'});
						$RowNo++;
					}
					$AssetsByProviderW = IOFactory::createWriter($AssetsByProviderSS, 'Xlsx');
					$AssetsByProviderW->save($EmbeddedAssetsByProvider);
				}

				// Assets By Location - Slide 11
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Assets by Location");
				$AssetsByLocation = $CubeJSResults['AssetsByLocation']['Body'];
				if (isset($AssetsByLocation->result->data)) {
					$EmbeddedAssetsByLocation = getEmbeddedSheetFilePath('AssetsByLocation', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$AssetsByLocationSS = IOFactory::load($EmbeddedAssetsByLocation);
					$RowNo = 2;
					$AssetsByLocationIncludedCount = 0;
					foreach ($AssetsByLocation->result->data as $AssetLocation) {
						$AssetsByLocationS = $AssetsByLocationSS->getActiveSheet();
						$AssetsByLocationS->setCellValue('A'.$RowNo, $this->convertProvider($AssetLocation->{'AssetDetails_ch.provider_label'}).': '.$AssetLocation->{'AssetDetails_ch.location_label'});
						$AssetsByLocationS->setCellValue('B'.$RowNo, round((100 / $TotalAssetsCount) * $AssetLocation->{'AssetDetails_ch.assetsCount'},2) / 100);
						$AssetsByLocationIncludedCount += $AssetLocation->{'AssetDetails_ch.assetsCount'};
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

				// Assets Missing from CMDB (Chart) - Slide 13
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Assets missing from CMDB");
				$AssetsMissingFromCMDB = $CubeJSResults['AssetsMissingFromCMDB']['Body'];
				if (isset($AssetsMissingFromCMDB->result->data)) {
					$EmbeddedAssetsMissingFromCMDB = getEmbeddedSheetFilePath('AssetsMissingFromCMDB', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$AssetsMissingFromCMDBSS = IOFactory::load($EmbeddedAssetsMissingFromCMDB);
					$RowNo = 2;
					foreach ($AssetsMissingFromCMDB->result->data as $AssetCategory) {
						$AssetsMissingFromCMDBS = $AssetsMissingFromCMDBSS->getActiveSheet();
						$AssetsMissingFromCMDBS->setCellValue('A'.$RowNo, $AssetCategory->{'AssetDetails_ch.providers_iterator'}); // Provider
						$AssetsMissingFromCMDBS->setCellValue('B'.$RowNo, $AssetCategory->{'AssetDetails_ch.assetsCount'});
						$RowNo++;
					}
					$AssetsMissingFromCMDBW = IOFactory::createWriter($AssetsMissingFromCMDBSS, 'Xlsx');
					$AssetsMissingFromCMDBW->save($EmbeddedAssetsMissingFromCMDB);
				}

				// Assets Missing from CMDB (Table) - Slide 13
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building top 3 Assets missing from CMDB");
				$Top3AssetsMissingFromCMDB = $CubeJSResults['Top3AssetsMissingFromCMDB']['Body'];
				if (isset($Top3AssetsMissingFromCMDB->result->data)) {
					$EmbeddedTop3AssetsMissingFromCMDB = getEmbeddedSheetFilePath('AssetsMissingFromCMDBTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$Top3AssetsMissingFromCMDBSS = IOFactory::load($EmbeddedTop3AssetsMissingFromCMDB);
					$RowNo = 2;
					foreach ($Top3AssetsMissingFromCMDB->result->data as $AssetCategory) {
						$Top3AssetsMissingFromCMDBS = $Top3AssetsMissingFromCMDBSS->getActiveSheet();
						$Top3AssetsMissingFromCMDBS->setCellValue('A'.$RowNo, $AssetCategory->{'AssetDetails_ch.name'}); // Asset Name
						$Top3AssetsMissingFromCMDBS->setCellValue('B'.$RowNo, $AssetCategory->{'AssetDetails_ch.location_label'}); // Asset Location
						$Top3AssetsMissingFromCMDBS->setCellValue('C'.$RowNo, $AssetCategory->{'AssetDetails_ch.taxonomy_type_label'}); // Asset Category / Type
						$Top3AssetsMissingFromCMDBS->setCellValue('D'.$RowNo, $AssetCategory->{'AssetDetails_ch.ip_address'}); // Asset IP Address
						$Top3AssetsMissingFromCMDBS->setCellValue('E'.$RowNo, $AssetCategory->{'AssetDetails_ch.provider_label'}); // Asset Provider
						$Top3AssetsMissingFromCMDBS->setCellValue('F'.$RowNo, 'Missing'); // Asset Status
						$RowNo++;
					}
					$Top3AssetsMissingFromCMDBW = IOFactory::createWriter($Top3AssetsMissingFromCMDBSS, 'Xlsx');
					$Top3AssetsMissingFromCMDBW->save($EmbeddedTop3AssetsMissingFromCMDB);
				}

				// Assets Orphaned in CMDB - Slide 13
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Assets orphaned in CMDB");
				$AssetsOrphanedInCMDB = $CubeJSResults['AssetsOrphanedInCMDB']['Body'];
				if (isset($AssetsOrphanedInCMDB->result->data)) {
					$EmbeddedAssetsOrphanedInCMDB = getEmbeddedSheetFilePath('AssetsOrphanedInCMDB', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$AssetsOrphanedInCMDBSS = IOFactory::load($EmbeddedAssetsOrphanedInCMDB);
					$RowNo = 2;
					foreach ($AssetsOrphanedInCMDB->result->data as $AssetCategory) {
						$AssetsOrphanedInCMDBS = $AssetsOrphanedInCMDBSS->getActiveSheet();
						$AssetsOrphanedInCMDBS->setCellValue('A'.$RowNo, $AssetCategory->{'AssetDetails_ch.providers_iterator'}); // Provider
						$AssetsOrphanedInCMDBS->setCellValue('B'.$RowNo, $AssetCategory->{'AssetDetails_ch.assetsCount'});
						$RowNo++;
					}
					$AssetsOrphanedInCMDBW = IOFactory::createWriter($AssetsOrphanedInCMDBSS, 'Xlsx');
					$AssetsOrphanedInCMDBW->save($EmbeddedAssetsOrphanedInCMDB);
				}
				// New Slides

				// Overlapping Subnets - Slide 16
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
						$OverlappingSubnetsS->setCellValue('C'.$RowNo, implode(', ', array_map([$this, 'convertProvider'], $ProviderType->{'NetworkInsightsOverlappingBlocksRolled.providers'})));
						$RowNo++;
					}
					$OverlappingSubnetsW = IOFactory::createWriter($OverlappingSubnetsSS, 'Xlsx');
					$OverlappingSubnetsW->save($EmbeddedOverlappingSubnets);
				}

				// Total IPs By Provider - Slide 16
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Total IPs by Provider");
				$EmbeddedTotalIPsByProvider = getEmbeddedSheetFilePath('TotalIPsByProvider', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$TotalIPsByProviderSS = IOFactory::load($EmbeddedTotalIPsByProvider);
				$TotalIPsByProviderS = $TotalIPsByProviderSS->getActiveSheet();
				$RowNo = 2;
				if ($GCPIPsCount > 0) {
					$TotalIPsByProviderS->setCellValue('A'.$RowNo, $this->convertProvider('GCP'));
					$TotalIPsByProviderS->setCellValue('B'.$RowNo, $GCPIPsCount);
					$RowNo++;
				}
				if ($AzureIPsCount > 0) {
					$TotalIPsByProviderS->setCellValue('A'.$RowNo, $this->convertProvider('Azure'));
					$TotalIPsByProviderS->setCellValue('B'.$RowNo, $AzureIPsCount);
					$RowNo++;
				}
				if ($AWSIPsCount > 0) {
					$TotalIPsByProviderS->setCellValue('A'.$RowNo, $this->convertProvider('AWS'));
					$TotalIPsByProviderS->setCellValue('B'.$RowNo, $AWSIPsCount);
					$RowNo++;
				}
				if ($OCIIPsCount > 0) {
					$TotalIPsByProviderS->setCellValue('A'.$RowNo, $this->convertProvider('Oracle Cloud Infrastructure'));
					$TotalIPsByProviderS->setCellValue('B'.$RowNo, $OCIIPsCount);
					$RowNo++;
				}
				$TotalIPsByProviderW = IOFactory::createWriter($TotalIPsByProviderSS, 'Xlsx');
				$TotalIPsByProviderW->save($EmbeddedTotalIPsByProvider);

				// Overutilized Subnets - Slide 18
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
						$OverutilizedSubnetsS->setCellValue('C'.$RowNo, $this->convertProvider($ProviderType->{'NetworkInsightsSubnet.provider'}));
						$OverutilizedSubnetsS->setCellValue('D'.$RowNo, round($ProviderType->{'NetworkInsightsSubnet.utilization_percent'},1) . '%');
						$RowNo++;
					}
					$OverutilizedSubnetsW = IOFactory::createWriter($OverutilizedSubnetsSS, 'Xlsx');
					$OverutilizedSubnetsW->save($EmbeddedOverutilizedSubnets);
				}
				
				// Underutilized Subnets - Slide 18
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
						$UnderutilizedSubnetsS->setCellValue('C'.$RowNo, $this->convertProvider($ProviderType->{'NetworkInsightsSubnet.provider'}));
						$UnderutilizedSubnetsS->setCellValue('D'.$RowNo, round($ProviderType->{'NetworkInsightsSubnet.utilization_percent'},1) . '%');
						$RowNo++;
					}
					$UnderutilizedSubnetsW = IOFactory::createWriter($UnderutilizedSubnetsSS, 'Xlsx');
					$UnderutilizedSubnetsW->save($EmbeddedUnderutilizedSubnets);
				}

				// High-Risk DNS Records by Category - Slide 21
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

				// Dangling Records - Slide 21
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Dangling Records");
				$Top3DanglingRecords = $CubeJSResults['Top3DanglingRecords']['Body'];
				if (isset($Top3DanglingRecords->result->data)) {
					$EmbeddedTop3DanglingRecords = getEmbeddedSheetFilePath('Top3DanglingRecordsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$Top3DanglingRecordsSS = IOFactory::load($EmbeddedTop3DanglingRecords);
					$RowNo = 2;
					foreach ($Top3DanglingRecords->result->data as $DanglingRecord) {
						$Top3DanglingRecordsS = $Top3DanglingRecordsSS->getActiveSheet();
						$Top3DanglingRecordsS->setCellValue('A'.$RowNo, $DanglingRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name'});
						$Top3DanglingRecordsS->setCellValue('B'.$RowNo, $DanglingRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name'});
						$Top3DanglingRecordsS->setCellValue('C'.$RowNo, $DanglingRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type'});
						$Top3DanglingRecordsS->setCellValue('D'.$RowNo, implode(', ',$DanglingRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids'}));
						$Top3DanglingRecordsS->setCellValue('E'.$RowNo, $this->convertProvider($DanglingRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str'}));
						$RowNo++;
					}
					$Top3DanglingRecordsW = IOFactory::createWriter($Top3DanglingRecordsSS, 'Xlsx');
					$Top3DanglingRecordsW->save($EmbeddedTop3DanglingRecords);
				}

				// Abandoned Records - Slide 22
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Abandoned Records");
				$Top3AbandonedRecords = $CubeJSResults['Top3AbandonedRecords']['Body'];
				if (isset($Top3AbandonedRecords->result->data)) {
					$EmbeddedTop3AbandonedRecords = getEmbeddedSheetFilePath('Top3AbandonedRecordsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$Top3AbandonedRecordsSS = IOFactory::load($EmbeddedTop3AbandonedRecords);
					$RowNo = 2;
					foreach ($Top3AbandonedRecords->result->data as $AbandonedRecord) {
						$Top3AbandonedRecordsS = $Top3AbandonedRecordsSS->getActiveSheet();
						$Top3AbandonedRecordsS->setCellValue('A'.$RowNo, $AbandonedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name'});
						$Top3AbandonedRecordsS->setCellValue('B'.$RowNo, $AbandonedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name'});
						$Top3AbandonedRecordsS->setCellValue('C'.$RowNo, $AbandonedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type'});
						$Top3AbandonedRecordsS->setCellValue('D'.$RowNo, implode(', ',$AbandonedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids'}));
						$Top3AbandonedRecordsS->setCellValue('E'.$RowNo, $this->convertProvider($AbandonedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str'}));
						$RowNo++;
					}
					$Top3AbandonedRecordsW = IOFactory::createWriter($Top3AbandonedRecordsSS, 'Xlsx');
					$Top3AbandonedRecordsW->save($EmbeddedTop3AbandonedRecords);
				}

				// Untrusted Records - Slide 22
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Untrusted Records");
				$Top3UntrustedRecords = $CubeJSResults['Top3UntrustedRecords']['Body'];
				if (isset($Top3UntrustedRecords->result->data)) {
					$EmbeddedTop3UntrustedRecords = getEmbeddedSheetFilePath('Top3UntrustedRecordsTable', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
					$Top3UntrustedRecordsSS = IOFactory::load($EmbeddedTop3UntrustedRecords);
					$RowNo = 2;
					foreach ($Top3UntrustedRecords->result->data as $UntrustedRecord) {
						$Top3UntrustedRecordsS = $Top3UntrustedRecordsSS->getActiveSheet();
						$Top3UntrustedRecordsS->setCellValue('A'.$RowNo, $UntrustedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_name'});
						$Top3UntrustedRecordsS->setCellValue('B'.$RowNo, $UntrustedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_absolute_zone_name'});
						$Top3UntrustedRecordsS->setCellValue('C'.$RowNo, $UntrustedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_record_type'});
						$Top3UntrustedRecordsS->setCellValue('D'.$RowNo, implode(', ',$UntrustedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.indicator_ids'}));
						$Top3UntrustedRecordsS->setCellValue('E'.$RowNo, $this->convertProvider($UntrustedRecord->{'NetworkInsightsDnsRecordsRolledByIndicatorId.record_asset_provider_types_str'}));
						$RowNo++;
					}
					$Top3UntrustedRecordsW = IOFactory::createWriter($Top3UntrustedRecordsSS, 'Xlsx');
					$Top3UntrustedRecordsW->save($EmbeddedTop3UntrustedRecords);
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
				rmdirRecursive($SelectedTemplate['ExtractedDir']);

				// Extract Powerpoint Template Strings
				// ** Using external library to save re-writing the string replacement functions manually. Will probably pull this in as native code at some point.
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Extract Powerpoint Strings");
				$extractor = new BasicExtractor();
				$mapping = $extractor->extractStringsAndCreateMappingFile(
					$this->getDir()['Files'].'/reports/report'.'-'.$config['UUID'].'-'.$SelectedTemplate['FileName'],
					$SelectedTemplate['ExtractedDir'].'-extracted'
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
				$mapping = replaceTag($mapping,'#TAG01',number_abbr($HighRiskAssetsCount)); // All High-Risk Assets
				$mapping = replaceTag($mapping,'#TAG02',number_abbr($AssetsMissingFromServiceNow)); // Inventory Gap Analysis - Missing Assets from CMDB - TAG02
				$mapping = replaceTag($mapping,'#TAG03',number_abbr($CloudSubnetOverlapCount)); // Cloud Subnet Overlap Count
				$mapping = replaceTag($mapping,'#TAG04',number_abbr($CloudSubnetUtilizationAbove50Total)); // Overutilized Subnets (>=50%)
				$mapping = replaceTag($mapping,'#TAG05',number_abbr($CloudSubnetUtilizationBelow50Total)); // Underutilized Subnets (<50%)
				$mapping = replaceTag($mapping,'#TAG06',number_abbr($DNSComplexityScore)); // DNS Complexity Score
				$mapping = replaceTag($mapping,'#TAG07',number_abbr($DanglingDNSCount)); // High-Risk DNS Records - Dangling
				$mapping = replaceTag($mapping,'#TAG08',number_abbr($AbandonedDNSCount)); // High-Risk DNS Records - Abandoned
				$mapping = replaceTag($mapping,'#TAG09',number_abbr($UntrustedDNSCount)); // High-Risk DNS Records - Untrusted
				
				##// Slide 7 - Cloud Assets
				$mapping = replaceTag($mapping,'#TAG10',number_abbr($TotalCloudAssetsCount)); // Total Assets
				$mapping = replaceTag($mapping,'#TAG11',number_abbr($HighRiskCloudAssetsCount)); // High-Risk Assets
				$mapping = replaceTag($mapping,'#TAG12',number_abbr($ZombieCloudAssetsCount)); // Zombie Assets Count
				$mapping = replaceTag($mapping,'#TAG13',number_abbr($UnregisteredCloudAssetsCount)); // Assets Missing Record(s)
				$mapping = replaceTag($mapping,'#TAG14',number_abbr($NonCompliantCloudAssetsCount)); // Non-Compliant Assets

				##// Slide 8 - Cloud Assets
				$mapping = replaceTag($mapping,'#TAG15',number_abbr($ZombieCloudAssetsCount)); // Zombie Assets Count
				$mapping = replaceTag($mapping,'#TAG16',number_abbr($ZombieCloudAssets['resourceutilizationidle'] ?? 0)); // Idle Zombie Assets Count
				$mapping = replaceTag($mapping,'#TAG17',$ZombieCloudAssetsResourceUtilizationIdlePerc); // Idle Zombie Assets Percentage
				$mapping = replaceTag($mapping,'#TAG18',number_abbr($ZombieCloudAssets['orphan'] ?? 0)); // Orphaned Assets Count
				$mapping = replaceTag($mapping,'#TAG19',$ZombieCloudAssetsOrphanedPerc); // Orphaned Assets Percentage

				##// Slide 9 - Cloud Assets
				$mapping = replaceTag($mapping,'#TAG20',number_abbr($UnregisteredCloudAssetsCount)); // Assets Missing Record(s)
				$mapping = replaceTag($mapping,'#TAG21',number_abbr($NonCompliantCloudAssetsCount)); // Non-Compliant Assets

				##// Slide 11 - Comprehensive Asset Inventory
				$mapping = replaceTag($mapping,'#TAG22',number_abbr($TotalAssetsCount)); // Total Assets

				##// Slide 12 - CMDB Assets
				$mapping = replaceTag($mapping,'#TAG23',number_abbr($AssetsMatchingInServiceNowPerc)); // Completeness - % of Assets found matching in CMDB - TAG23
				$mapping = replaceTag($mapping,'#TAG24',number_abbr($AssetsOnlyInServiceNow)); // ServiceNow / Total - Total number of assets found in ServiceNow Only - TAG24
				$mapping = replaceTag($mapping,'#TAG25',number_abbr($AssetsOnlyInServiceNowPerc)); // ServiceNow / Total - Total percentage of assets found in ServiceNow Only - TAG25
				$mapping = replaceTag($mapping,'#TAG26',number_abbr($AssetsMatchingInServiceNow)); // Overlap - Total Matching Assets between Infoblox and ServiceNow - TAG26
				$mapping = replaceTag($mapping,'#TAG27',number_abbr($AssetsMissingFromServiceNow)); // Infoblox / Total - Total number of assets found in Infoblox Only - TAG27
				$mapping = replaceTag($mapping,'#TAG28',number_abbr($AssetsMissingFromServiceNowPerc)); // Infoblox / Total - Total percentage of assets found in Infoblox Only - TAG28
				$mapping = replaceTag($mapping,'#TAG29',number_abbr($TotalAssetsCount)); // Total CMDB Assets
				$mapping = replaceTag($mapping,'#TAG30',number_abbr($AssetsOnlyInServiceNow+$AssetsMissingFromServiceNow)); // Discrepencies (Total number of assets only in SNOW + total number of assets not in SNOW) - TAG30
				$mapping = replaceTag($mapping,'#TAG31',number_abbr($AssetsOnlyInServiceNow)); // Only in ServiceNow - TAG31
				$mapping = replaceTag($mapping,'#TAG32',number_abbr($AssetsMatchingInServiceNow)); // Matching - TAG32 (Same as TAG26)
				$mapping = replaceTag($mapping,'#TAG33',number_abbr($AssetsMissingFromServiceNow)); // Missing from ServiceNow - TAG33
							

				##// Slide 13 - CMDB Assets
				$mapping = replaceTag($mapping,'#TAG34',number_abbr($AssetsMissingFromServiceNow)); // Infoblox / Total - Total number of assets found in Infoblox Only - TAG34
				$mapping = replaceTag($mapping,'#TAG35',number_abbr($AssetsOnlyInServiceNow)); // Infoblox / Total - Total number of assets found in ServiceNow Only - TAG35

				##// Slide 14 - CMDB Assets (Loop per Asset Provider) - TO DO LATER
				// #PV01 - Provider Name
				// #PV02 - Provider Name
				// #PV03 - Provider Name
				// #PV04 - Provider Name
				// #PV05 - Provider Name
				// #PV06 - Total Assets for Provider
				// #PV07 - Provider Name
				// #PV08 - Total Assets unique to Provider
				// #PV09 - Total Assets unique to Provider Percentage
				// #PV10 - Total Assets in other Providers
				// #PV11 - Total Assets in other Providers Percentage

				##// Slide 16 - IP/Subnet Allocation
				$mapping = replaceTag($mapping,'#TAG36',number_abbr($TotalCloudSubnetsCount)); // Total Subnets (Box)
				$mapping = replaceTag($mapping,'#TAG37',number_abbr($TotalCloudSubnetsCount)); // Total Subnets (Centre)
				$mapping = replaceTag($mapping,'#TAG38',number_abbr($AWSSubnetsCount)); // AWS Subnet Count
				$mapping = replaceTag($mapping,'#TAG39',number_abbr($AWSSubnetsPercentage)); // AWS Subnet Percentage
				$mapping = replaceTag($mapping,'#TAG40',number_abbr($AzureSubnetsCount)); // Azure Subnet Count
				$mapping = replaceTag($mapping,'#TAG41',number_abbr($AzureSubnetsPercentage)); // Azure Subnet Percentage
				$mapping = replaceTag($mapping,'#TAG42',number_abbr($GCPSubnetsCount)); // GCP Subnet Count
				$mapping = replaceTag($mapping,'#TAG43',number_abbr($GCPSubnetsPercentage)); // GCP Subnet Percentage
				$mapping = replaceTag($mapping,'#TAG44',number_abbr($OCISubnetsCount)); // OCI Subnet Count
				$mapping = replaceTag($mapping,'#TAG45',number_abbr($OCISubnetsPercentage)); // OCI Subnet Percentage
				
				$mapping = replaceTag($mapping,'#TAG46',number_abbr($TotalCloudIPsCount)); // Total Allocated Cloud IPs
				$mapping = replaceTag($mapping,'#TAG47',number_abbr($AzureIPsCount)); // Total Allocated Azure IPs
				$mapping = replaceTag($mapping,'#TAG48',number_abbr($AWSIPsCount)); // Total Allocated AWS IPs
				$mapping = replaceTag($mapping,'#TAG49',number_abbr($GCPIPsCount)); // Total Allocated GCP IPs
				$mapping = replaceTag($mapping,'#TAG50',number_abbr($OCIIPsCount)); // Total Allocated OCI IPs

				##// Slide 17 - IP/Subnet Allocation
				$mapping = replaceTag($mapping,'#TAG51',number_abbr($AWSSubnetsCount)); // Total AWS Subnets
				$mapping = replaceTag($mapping,'#TAG52',number_abbr($CloudSubnetUtilizationAbove50ByProvider['AWS'])); // AWS Subnets above 50% Utilization
				$mapping = replaceTag($mapping,'#TAG53',number_abbr($CloudSubnetUtilizationBelow50ByProvider['AWS'])); // AWS Subnets below 50% Utilization
				$mapping = replaceTag($mapping,'#TAG54',number_abbr($AzureSubnetsCount)); // Total Azure Subnets
				$mapping = replaceTag($mapping,'#TAG55',number_abbr($CloudSubnetUtilizationAbove50ByProvider['Azure'])); // Azure Subnets above 50% Utilization
				$mapping = replaceTag($mapping,'#TAG56',number_abbr($CloudSubnetUtilizationBelow50ByProvider['Azure'])); // Azure Subnets below 50% Utilization
				$mapping = replaceTag($mapping,'#TAG57',number_abbr($GCPSubnetsCount)); // Total GCP Subnets
				$mapping = replaceTag($mapping,'#TAG58',number_abbr($CloudSubnetUtilizationAbove50ByProvider['GCP'])); // GCP Subnets above 50% Utilization
				$mapping = replaceTag($mapping,'#TAG59',number_abbr($CloudSubnetUtilizationBelow50ByProvider['GCP'])); // GCP Subnets below 50% Utilization
				$mapping = replaceTag($mapping,'#TAG60',number_abbr($OCISubnetsCount)); // Total OCI Subnets
				$mapping = replaceTag($mapping,'#TAG61',number_abbr($CloudSubnetUtilizationAbove50ByProvider['Oracle Cloud Infrastructure'])); // OCI Subnets above 50% Utilization
				$mapping = replaceTag($mapping,'#TAG62',number_abbr($CloudSubnetUtilizationBelow50ByProvider['Oracle Cloud Infrastructure'])); // OCI Subnets below 50% Utilization
				
				##// Slide 20 - DNS Complexity
				$mapping = replaceTag($mapping,'#TAG63',number_abbr($OverlappingZonesCount)); // Overlapping DNS Zones
				$mapping = replaceTag($mapping,'#TAG64',number_abbr($DNSComplexityScore)); // DNS Complexity Score
				$mapping = replaceTag($mapping,'#TAG65',number_abbr($CloudDNSZonesByProviderTotals['Azure'])); // Azure Zones
				$mapping = replaceTag($mapping,'#TAG66',number_abbr($CloudDNSZonesByProviderTotals['AWS'])); // AWS Zones
				$mapping = replaceTag($mapping,'#TAG67',number_abbr($CloudDNSZonesByProviderTotals['GCP'])); // GCP Zones
				$mapping = replaceTag($mapping,'#TAG68',number_abbr($CloudDNSZonesByProviderTotals['Oracle Cloud Infrastructure'])); // OCI Zones
				$mapping = replaceTag($mapping,'#TAG69',number_abbr($CloudDNSRecordsByProviderTotals['microsoft_azure'])); // Azure Records
				$mapping = replaceTag($mapping,'#TAG70',number_abbr($CloudDNSRecordsByProviderTotals['amazon_web_service'])); // AWS Records
				$mapping = replaceTag($mapping,'#TAG71',number_abbr($CloudDNSRecordsByProviderTotals['google_cloud_platform'])); // GCP Records
				$mapping = replaceTag($mapping,'#TAG72',number_abbr($CloudDNSRecordsByProviderTotals['oracle_cloud_infrastructure'])); // OCI Records
				$mapping = replaceTag($mapping,'#TAG73',number_abbr($CloudForwarderCounts['Microsoft Azure']['inbound'])); // Azure Inbound Endpoints
				$mapping = replaceTag($mapping,'#TAG74',number_abbr($CloudForwarderCounts['Microsoft Azure']['outbound'])); // Azure Outbound Endpoints
				$mapping = replaceTag($mapping,'#TAG75',number_abbr($CloudForwarderCounts['Amazon Web Services']['inbound'])); // AWS Inbound Endpoints
				$mapping = replaceTag($mapping,'#TAG76',number_abbr($CloudForwarderCounts['Amazon Web Services']['outbound'])); // AWS Outbound Endpoints
				$mapping = replaceTag($mapping,'#TAG77',number_abbr($CloudForwarderCounts['Google Cloud Platform']['inbound'])); // GCP Inbound Endpoints
				$mapping = replaceTag($mapping,'#TAG78',number_abbr($CloudForwarderCounts['Google Cloud Platform']['outbound'])); // GCP Outbound Endpoints
				$mapping = replaceTag($mapping,'#TAG79',number_abbr($CloudForwarderCounts['Oracle Cloud Infrastructure']['inbound'])); // OCI Inbound Endpoints
				$mapping = replaceTag($mapping,'#TAG80',number_abbr($CloudForwarderCounts['Oracle Cloud Infrastructure']['outbound'])); // OCI Outbound Endpoints

				##// Slide 21 - High-Risk DNS Records
				$mapping = replaceTag($mapping,'#TAG81',number_abbr($HighRiskDNSRecordsCount)); // High-Risk DNS Records
				$mapping = replaceTag($mapping,'#TAG82',number_abbr($DanglingDNSCount)); // High-Risk DNS Records - Dangling
				$mapping = replaceTag($mapping,'#TAG83',number_abbr($AbandonedDNSCount)); // High-Risk DNS Records - Abandoned
				$mapping = replaceTag($mapping,'#TAG84',number_abbr($UntrustedDNSCount)); // High-Risk DNS Records - Untrusted

				##// Slide 20 - Recommendations
				$mapping = replaceTag($mapping,'#TAG85',number_abbr($TokensManagementCount)); // Management Tokens
				$mapping = replaceTag($mapping,'#TAG86',number_abbr($TokensActiveIPPercentage)); // Active IP Percentage
				$mapping = replaceTag($mapping,'#TAG87',number_abbr($ActiveIPCount)); // Active IP Count
				$mapping = replaceTag($mapping,'#TAG88',number_abbr($TokensAssetPercentage)); // Asset Count
				$mapping = replaceTag($mapping,'#TAG89',number_abbr($TokensDDIPercentage)); // DDI Objects
				$mapping = replaceTag($mapping,'#TAG90',number_abbr($TokensServer)); // Server Tokens
				$mapping = replaceTag($mapping,'#TAG91',number_abbr($TokensReporting)); // Reporting Tokens

				// Rebuild Powerpoint File(s)
				// ** Using external library to save re-writing the string replacement functions manually. Will probably pull this in as native code at some point.
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Rebuilding Powerpoint Template(s)");
				$injector = new BasicInjector();
				$injector->injectMappingAndCreateNewFile(
					$mapping,
					$SelectedTemplate['ExtractedDir'].'-extracted',
					$SelectedTemplate['ExtractedDir'].$SelectedTemplateFileExt
				);
		
				// Cleanup
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Final Cleanup");
				unlink($SelectedTemplate['ExtractedDir'].'-extracted');
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
			'Total' => ($Total * 28) + 18,
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

	public function convertProvider($Provider) {
		switch ($Provider) {
			case 'AWS':
			case 'amazon_web_service':
				return 'AWS';
			case 'Azure':
			case 'microsoft_azure':
				return 'Azure';
			case 'GCP':
			case 'google_cloud_platform':
				return 'Google Cloud';
			case 'OCI':
			case 'oracle_cloud_infrastructure':
				return 'Oracle Cloud Infrastructure';
			default:
				return $Provider;
		}
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