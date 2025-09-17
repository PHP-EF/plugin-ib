<?php

use Label305\PptxExtractor\Basic\BasicExtractor;
use Label305\PptxExtractor\Basic\BasicInjector;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class SecurityAssessment extends ibPortal {
	private $AssessmentReporting;
	private $TemplateConfig;
	private $ThreatActors;

	public function __construct() {
		parent::__construct();
		if (!is_dir($this->getDir()['Files'].'/reports')) {
			mkdir($this->getDir()['Files'].'/reports', 0755, true);
		}
		$this->AssessmentReporting = new AssessmentReporting();
		$this->TemplateConfig = new TemplateConfig();
	}

	public function generateSecurityReport($config) {
		// Pass APIKey & Realm to ThreatActors Class
		$this->ThreatActors = new ThreatActors();
		$this->ThreatActors->SetCSPConfiguration($this->APIKey,$config['Realm']);
		// Check Active Template Exists
		$SelectedTemplates = [];
		foreach ($config['Templates'] as $TemplateId) {
			$SelectedTemplate = $this->TemplateConfig->getSecurityAssessmentTemplateConfigById($TemplateId);
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
			$this->logging->writeLog("Assessment",$UserInfo->result->name." requested a security assessment report for: ".$CurrentAccount->name,"info");
			$ReportRecordId = $this->AssessmentReporting->newReportEntry('Security Assessment',$UserInfo->result->name,$CurrentAccount->name,$config['Realm'],$config['UUID'],"Started");
	
			// Set Progress
			$Progress = 0;
	
			// Set Time Dimensions
			$StartDimension = str_replace('Z','',$config['StartDateTime']);
			$EndDimension = str_replace('Z','',$config['EndDateTime']);
	
			$HighRiskCategoryList = implode('","',[
				"Risky Activity",
				"Suspicious and Malicious Software",
				// "Uncategorized",
				"Adult",
				"Abortion",
				"Abortion Pro Choice",
				"Abortion Pro Life",
				"Child Inappropriate",
				"Gambling",
				"Gay",
				"Lingerie",
				"Nudity",
				"Pornography",
				"Profanity",
				"R-Rated",
				"Sex & Erotic",
				"Sex Education",
				"Tobacco",
				"Anonymizer",
				"Criminal Skills",
				"Self Harm",
				"Criminal Activities - Other",
				"Illegal Drugs",
				"Marijuana",
				"Child Abuse Images",
				"Hacking",
				"Hate Speech",
				"Piracy & Copyright Theft",
				"Torrent Repository",
				"Terrorism",
				"Peer-to-Peer",
				"Violence",
				"Weapons",
				"School Cheating",
				"Ad Fraud",
				"Botnet",
				"Command and Control Centers",
				"Compromised & Links To Malware",
				"Malware Call-Home",
				"Malware Distribution Point",
				"Phishing/Fraud",
				"Spam URLs",
				"Spyware & Questionable Software",
				"Cryptocurrency Mining",
				"Sexuality",
				"Parked & For Sale Domains"
			]);
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Collecting Metrics");
			$CubeJSRequests = array(
				'TopThreatFeeds' => '{"measures":["PortunusAggSecurity.feednameCount"],"dimensions":["PortunusAggSecurity.feed_name"],"timeDimensions":[{"dimension":"PortunusAggSecurity.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggSecurity.type","operator":"equals","values":["2"]},{"member":"PortunusAggSecurity.severity","operator":"equals","values":["High"]}],"limit":"10","ungrouped":false}',
				'TopDetectedProperties' => '{"measures":["PortunusDnsLogs.tpropertyCount"],"dimensions":["PortunusDnsLogs.tproperty"],"timeDimensions":[{"dimension":"PortunusDnsLogs.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusDnsLogs.type","operator":"equals","values":["2"]},{"member":"PortunusDnsLogs.feed_name","operator":"notEquals","values":["Public_DOH","public-doh","Public_DOH_IP","public-doh-ip"]},{"member":"PortunusDnsLogs.severity","operator":"notEquals","values":["Low","Info"]}],"limit":"10","ungrouped":false}',
				// Switch to using the Web Content Discovery APIs, rather than those called via the Dashboard.
				//'ContentFiltration' => '{"measures":["PortunusAggWebcontent.categoryCount"],"dimensions":["PortunusAggWebcontent.category"],"timeDimensions":[{"dimension":"PortunusAggWebcontent.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[],"limit":"10","ungrouped":false}',
				// Switch to using data from High-Risk websites instead, this is now an unneccessary API call
				//'ContentFiltration' => '{"timeDimensions":[{"dimension":"PortunusAggWebContentDiscovery.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"measures":["PortunusAggWebContentDiscovery.count"],"dimensions":["PortunusAggWebContentDiscovery.domain_category"],"filters":[{"member":"PortunusAggWebContentDiscovery.domain_category","operator":"set"},{"member":"PortunusAggWebContentDiscovery.domain_category","operator":"notEquals","values":[""]}],"order":{"PortunusAggWebContentDiscovery.count":"desc"},"limit":"10"}',
				'InsightDistribution' => '{"measures":["InsightsAggregated.count"],"dimensions":["InsightsAggregated.threatType"],"filters":[{"member":"InsightsAggregated.insightStatus","operator":"equals","values":["Active"]}]}',
				'DNSFirewallActivity' => '{"measures":["PortunusAggSecurity.severityCount"],"dimensions":["PortunusAggSecurity.severity"],"timeDimensions":[{"dimension":"PortunusAggSecurity.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggSecurity.type","operator":"equals","values":["2","3"]},{"member":"PortunusAggSecurity.severity","operator":"equals","values":["High","Medium","Low"]}],"limit":"3","ungrouped":false}',
				'DNSFirewallActivityDaily' => '{"measures":["PortunusAggSecurity.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggSecurity.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":"day"}],"filters":[{"member":"PortunusAggSecurity.type","operator":"equals","values":["2","3"]}],"order":{"PortunusAggSecurity.timestamp":"asc"}}',
				'DNSActivity' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["1"]}],"limit":"1","ungrouped":false}',
				'DNSActivityDaily' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":"day"}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["1"]}],"order":{"PortunusAggInsight.timestamp":"asc"}}',
				'SOCInsights' => '{"measures":["InsightsAggregated.count","InsightsAggregated.mostRecentAt","InsightsAggregated.startedAtMin"],"dimensions":["InsightsAggregated.priorityText"],"filters":[{"member":"InsightsAggregated.insightStatus","operator":"equals","values":["Active"]}],"timezone":"UTC"}',
				'SecurityEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"contains","values":["2","3"]}],"limit":"1","ungrouped":false}',
				'DataExfilEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["4"]},{"member":"PortunusAggInsight.tclass","operator":"equals","values":["TI-DNST"]}],"ungrouped":false}',
				// May need a review based on recent findings of missing data?
				'ZeroDayDNSEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2","3"]},{"member":"PortunusAggInsight.tclass","operator":"equals","values":["Zero Day DNS"]}],"ungrouped":false}',
				'SuspiciousEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2"]},{"member":"PortunusAggInsight.tclass","operator":"equals","values":["Suspicious"]}],"ungrouped":false}',
				'HighRiskWebsites' => '{"timeDimensions":[{"dimension":"PortunusAggWebContentDiscovery.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"measures":["PortunusAggWebContentDiscovery.count","PortunusAggWebContentDiscovery.deviceCount"],"dimensions":["PortunusAggWebContentDiscovery.domain_category"],"order":{"PortunusAggWebContentDiscovery.count":"desc"},"filters":[{"member":"PortunusAggWebContentDiscovery.domain_category","operator":"equals","values":["'.$HighRiskCategoryList.'"]}]}',
				'DOHEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2"]},{"member":"PortunusAggInsight.tproperty","operator":"equals","values":["DoHService"]}],"ungrouped":false}',
				'NODEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2"]},{"member":"PortunusAggInsight.tproperty","operator":"equals","values":["NewlyObservedDomains"]}],"ungrouped":false}',
				'DGAEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"or":[{"member":"PortunusAggInsight.tproperty","operator":"equals","values":["suspicious_rdga","suspicious_dga","DGA"]},{"member":"PortunusAggInsight.tclass","operator":"equals","values":["DGA","MalwareC2DGA"]}]},{"member":"PortunusAggInsight.type","operator":"equals","values":["2","3"]}],"ungrouped":false}',
				'AppDiscoveryApplications' => '{"measures":["PortunusAggAppDiscovery.requests","PortunusAggAppDiscovery.deviceCount"],"dimensions":["PortunusAggAppDiscovery.app_name","PortunusAggAppDiscovery.app_category","PortunusAggAppDiscovery.app_vendor","PortunusAggAppDiscovery.app_approval"],"timeDimensions":[{"dimension":"PortunusAggAppDiscovery.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggAppDiscovery.app_name","operator":"set"},{"member":"PortunusAggAppDiscovery.app_name","operator":"notEquals","values":[""]}],"order":{}}',
				'AppDiscoveryTotals' => '{"timeDimensions":[{"dimension":"PortunusAggAppDiscovery.timestamp","granularity":null,"dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"segments":[],"dimensions":[],"ungrouped":false,"measures":["PortunusAggAppDiscovery.requests","PortunusAggAppDiscovery.deviceCount"]}',
				'WebContentTotals' => '{"ungrouped":false,"measures":["PortunusAggWebContentDiscovery.deviceCount","PortunusAggWebContentDiscovery.requests"],"segments":[],"dimensions":[],"filters":[{"member":"PortunusAggWebContentDiscovery.domain_category","operator":"notEquals","values":[null]}],"timeDimensions":[{"dimension":"PortunusAggWebContentDiscovery.timestamp","granularity":null,"dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}]}',
				'WebContentSecurityEvents' => '{"measures":["PortunusAggWebcontent.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggWebcontent.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggWebcontent.type","operator":"equals","values":["3"]},{"member":"PortunusAggWebcontent.category","operator":"notEquals","values":[null]}],"limit":"1","ungrouped":false}',
				'ThreatActivityEvents' => '{"measures":["PortunusAggInsight.threatCount"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2"]},{"member":"PortunusAggInsight.severity","operator":"equals","values":["High","Medium","Low"]},{"member":"PortunusAggInsight.threat_indicator","operator":"notEquals","values":[""]}],"limit":"1","ungrouped":false}',
				'DNSFirewallEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"and":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2"]},{"or":[{"member":"PortunusAggInsight.severity","operator":"equals","values":["High","Medium","Low"]},{"and":[{"member":"PortunusAggInsight.severity","operator":"equals","values":["Info"]},{"member":"PortunusAggInsight.policy_action","operator":"equals","values":["Block","Log"]}]}]},{"member":"PortunusAggInsight.confidence","operator":"equals","values":["High","Medium","Low"]}]}],"limit":"1","ungrouped":false}',
				// This is the number of "impacted devices", whilst the next one is all devices
				// 'Devices' => '{"measures":["PortunusAggInsight.deviceCount"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"contains","values":["2","3"]},{"member":"PortunusAggInsight.severity","operator":"contains","values":["High","Medium","Low"]}],"limit":"1","ungrouped":false}',
				'Devices' => '{"measures":["PortunusAggInsight.deviceCount"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["1"]}],"limit":"1","ungrouped":false}',
				'Users' => '{"measures":["PortunusAggInsight.userCount"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"contains","values":["2","3"]}],"limit":"1","ungrouped":false}',
				'ThreatInsight' => '{"measures":[],"dimensions":["PortunusDnsLogs.tproperty"],"timeDimensions":[{"dimension":"PortunusDnsLogs.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusDnsLogs.type","operator":"equals","values":["4"]}],"limit":"10000","ungrouped":false}',
				'ThreatView' => '{"measures":["PortunusAggInsight.tpropertyCount"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2"]}],"limit":"1","ungrouped":false}',
				'Sources' => '{"measures":["PortunusAggSecurity.networkCount"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggSecurity.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggSecurity.type","operator":"contains","values":["2","3"]}],"limit":"1","ungrouped":false}',
				// Workaround for removal of batch threat actor enrichment
				//'ThreatActors' => '{"segments":[],"timeDimensions":[{"dimension":"PortunusAggIPSummary.timestamp","granularity":null,"dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"ungrouped":false,"order":{"PortunusAggIPSummary.timestampMax":"desc"},"measures":["PortunusAggIPSummary.count"],"dimensions":["PortunusAggIPSummary.threat_indicator","PortunusAggIPSummary.actor_id"],"limit":1000,"filters":[{"and":[{"operator":"set","member":"PortunusAggIPSummary.threat_indicator"},{"operator":"set","member":"PortunusAggIPSummary.actor_id"}]}]}'
				// Removed due to workaround below
				//'ThreatActors' => '{"measures":[],"segments":[],"dimensions":["ThreatActors.storageid","ThreatActors.ikbactorid","ThreatActors.domain","ThreatActors.ikbfirstsubmittedts","ThreatActors.vtfirstdetectedts","ThreatActors.firstdetectedts","ThreatActors.lastdetectedts"],"timeDimensions":[{"dimension":"ThreatActors.lastdetectedts","granularity":null,"dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"ungrouped":false}'
				'FirstToDetect' => '{"measures":["TIDERPZStatsDetails.domainCount","TIDERPZStatsDetails.minutesAheadOfIndustryAvg"],"timeDimensions":[{"dimension":"TIDERPZStatsDetails.lastSeenForCustomer","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["TIDERPZStatsDetails.threatClass"],"filters":[],"timezone":"UTC","segments":[]}',
				'BandwidthSavings' => '{"measures":["PortunusAggThreat_ch.bandwidthTotal"],"timeDimensions":[{"dimension":"PortunusAggThreat_ch.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["PortunusAggThreat_ch.tclass"],"filters":[{"member":"PortunusAggThreat_ch.action","operator":"equals","values":["Block"]}],"timezone":"UTC","segments":[]}',
				'BandwidthSavingsPercentage' => '{"measures":["PortunusAggThreat_ch.bandwidthSavedPercentage"],"timeDimensions":[{"dimension":"PortunusAggThreat_ch.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"dimensions":["PortunusAggThreat_ch.action"],"filters":[{"member":"PortunusAggThreat_ch.action","operator":"equals","values":["Block"]}],"timezone":"UTC","segments":[]}',
				'MaliciousTDS' => '{"measures":["PortunusAggThreat_ch.requests","PortunusAggThreat_ch.timestampMax","PortunusAggThreat_ch.totalAssetCount","PortunusAggThreat_ch.threatCount"],"dimensions":["PortunusAggThreat_ch.severity"],"filters":[{"member":"PortunusAggThreat_ch.severity","operator":"notEquals","values":["Info"]},{"member":"PortunusAggThreat_ch.tclass","operator":"equals","values":["Malicious"]},{"member":"PortunusAggThreat_ch.tproperty","operator":"equals","values":["TDS"]}],"segments":["PortunusAggThreat_ch.threat_classes"],"timeDimensions":[{"dimension":"PortunusAggThreat_ch.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"limit":39,"offset":0,"total":true,"order":{"PortunusAggThreat_ch.timestampMax":"asc"}}'
			);
			// Workaround for EU / US Realm Alignment
			// if ($config['Realm'] == 'EU') {
				// $CubeJSRequests['ThreatActors'] = '{"segments":[],"timeDimensions":[{"dimension":"PortunusAggIPSummary.timestamp","granularity":null,"dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"ungrouped":false,"order":{"PortunusAggIPSummary.timestampMax":"desc"},"measures":["PortunusAggIPSummary.count"],"dimensions":["PortunusAggIPSummary.threat_indicator","PortunusAggIPSummary.actor_id"],"limit":1000,"filters":[{"and":[{"operator":"set","member":"PortunusAggIPSummary.threat_indicator"},{"operator":"set","member":"PortunusAggIPSummary.actor_id"}]}]}';
			// } elseif ($config['Realm'] == 'US') {
				$CubeJSRequests['ThreatActors'] = '{"measures":[],"segments":[],"dimensions":["ThreatActors.storageid","ThreatActors.ikbactorid","ThreatActors.domain","ThreatActors.ikbfirstsubmittedts","ThreatActors.vtfirstdetectedts","ThreatActors.firstdetectedts","ThreatActors.lastdetectedts"],"timeDimensions":[{"dimension":"ThreatActors.lastdetectedts","granularity":null,"dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"ungrouped":false}';
			// }
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
				'TopDetectedProperties' => 5,
				'ContentFiltration' => 6,
				'DNSActivity' => 7,
				'DNSFirewallActivity' => 8,
				'InsightDistribution' => 9,
				'AppDiscovery' => 12,
				'WebContentDiscovery' => 13,
				'Lookalikes' => 14
			];

			// 10 - Outlier Insight (Assets)
			// 11 - Outlier Insight (Indicators/Events)

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

			// DNS Firewall Activity - Used on Slides 2, 5 & 6
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building DNS Firewall Event Criticality");
			$DNSFirewallActivity = $CubeJSResults['DNSFirewallActivity']['Body'];
			if (isset($DNSFirewallActivity->result)) {
				$HighId = array_search('High', array_column($DNSFirewallActivity->result->data, 'PortunusAggSecurity.severity'));
				$MediumId = array_search('Medium', array_column($DNSFirewallActivity->result->data, 'PortunusAggSecurity.severity'));
				$LowId = array_search('Low', array_column($DNSFirewallActivity->result->data, 'PortunusAggSecurity.severity'));
				if ($HighId !== false) {$HighEventsCount = $DNSFirewallActivity->result->data[$HighId]->{'PortunusAggSecurity.severityCount'};} else {$HighEventsCount = 0;}
				if ($MediumId !== false) {$MediumEventsCount = $DNSFirewallActivity->result->data[$MediumId]->{'PortunusAggSecurity.severityCount'};} else {$MediumEventsCount = 0;}
				if ($LowId !== false) {$LowEventsCount = $DNSFirewallActivity->result->data[$LowId]->{'PortunusAggSecurity.severityCount'};} else {$LowEventsCount = 0;}
			} else {
				$HighEventsCount = 0;
				$MediumEventsCount = 0;
				$LowEventsCount = 0;
			}
	
			$HML = $HighEventsCount+$MediumEventsCount+$LowEventsCount;
			if ($HML > 0) {
				$HMLP = 100 / $HML;
			} else {
				$HMLP = 0;
			}
			$HighPerc = $HighEventsCount * $HMLP;
			$MediumPerc = $MediumEventsCount * $HMLP;
			$LowPerc = $LowEventsCount * $HMLP;
	
			// Total DNS Activity - Used on Slides 6 & 9
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building DNS Activity");
			$DNSActivity = $CubeJSResults['DNSActivity']['Body'];
			if (isset($DNSActivity->result->data[0])) {
				$DNSActivityCount = $DNSActivity->result->data[0]->{'PortunusAggInsight.requests'};
			} else {
				$DNSActivityCount = 0;
			}

			// Bandwidth Savings - Used on Slides 5 & 7
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Bandwidth Savings");
			$BandwidthSavings = $CubeJSResults['BandwidthSavings']['Body'];
			$TotalBandwidthBytes = 0;
			if (isset($BandwidthSavings->result->data) && is_array($BandwidthSavings->result->data)) {
				foreach ($BandwidthSavings->result->data as $Entry) {
					$TotalBandwidthBytes += $Entry->{'PortunusAggThreat_ch.bandwidthTotal'};
				}
			}
			$TotalBandwidthSaved = $this->formatBytes($TotalBandwidthBytes);

			// Bandwidth Savings Percentage - Used on Slide 7
			$BandwidthSavingsPercentage = $CubeJSResults['BandwidthSavingsPercentage']['Body'];
			if (isset($BandwidthSavingsPercentage->result->data[0])) {
				$BandwidthSavedPercentage = round($BandwidthSavingsPercentage->result->data[0]->{'PortunusAggThreat_ch.bandwidthSavedPercentage'},2);
			} else {
				$BandwidthSavedPercentage = 0;
			}

			// Malicious TDS Domains - Used on Slide 6
			$MaliciousTDSCounts = [
				'Requests' => 0,
				'Domains' => 0,
			];
			$MaliciousTDS = $CubeJSResults['MaliciousTDS']['Body'];
			if (isset($MaliciousTDS->result->data) && is_array($MaliciousTDS->result->data)) {
				foreach ($MaliciousTDS->result->data as $Entry) {
					$MaliciousTDSCounts['Requests'] += $Entry->{'PortunusAggThreat_ch.requests'};
					$MaliciousTDSCounts['Domains'] += $Entry->{'PortunusAggThreat_ch.threatCount'};
				}
			}

			// First to Detect - Used on Slide 5
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building First to Detect");
			$FirstToDetect = $CubeJSResults['FirstToDetect']['Body'];
			$FirstToDetectData = [];
			if (isset($FirstToDetect->result->data) && is_array($FirstToDetect->result->data)) {
				foreach ($FirstToDetect->result->data as $Entry) {
					$Class = $Entry->{'TIDERPZStatsDetails.threatClass'};
					$FirstToDetectData[$Class] = [
						'DomainCount' => $Entry->{'TIDERPZStatsDetails.domainCount'},
						'MinutesAheadOfIndustryAvg' => $Entry->{'TIDERPZStatsDetails.minutesAheadOfIndustryAvg'}
					];
				}
			}
			$FirstToDetectTotalDomains = array_sum(array_column($FirstToDetectData, 'DomainCount'));

			// Lookalike Domains - Used on Slides 5, 6 & 24
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Getting Lookalike Domain Counts");
			$LookalikeDomainCounts = $this->QueryCSP("get","api/atcfw/v1/lookalike_domain_counts");
			if (isset($LookalikeDomainCounts->results->count_total)) { $LookalikeTotalCount = $LookalikeDomainCounts->results->count_total; } else { $LookalikeTotalCount = 0; }
			if (isset($LookalikeDomainCounts->results->percentage_increase_total)) { $LookalikeTotalPercentage = $LookalikeDomainCounts->results->percentage_increase_total; } else { $LookalikeTotalPercentage = 0; }
			if (isset($LookalikeDomainCounts->results->count_custom)) { $LookalikeCustomCount = $LookalikeDomainCounts->results->count_custom; } else { $LookalikeCustomCount = 0; }
			if (isset($LookalikeDomainCounts->results->percentage_increase_custom)) { $LookalikeCustomPercentage = $LookalikeDomainCounts->results->percentage_increase_custom; } else { $LookalikeCustomPercentage = 0; }
			if (isset($LookalikeDomainCounts->results->count_threats)) { $LookalikeThreatCount = $LookalikeDomainCounts->results->count_threats; } else { $LookalikeThreatCount = 0; }
			if (isset($LookalikeDomainCounts->results->percentage_increase_threats)) { $LookalikeThreatPercentage = $LookalikeDomainCounts->results->percentage_increase_threats; } else { $LookalikeThreatPercentage = 0; }
	
			// SOC Insights - Used on Slides 15 & 28
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building SOC Insight Threat Criticality");
			$SOCInsights = $CubeJSResults['SOCInsights']['Body'];
			if (isset($SOCInsights->result)) {
				$InfoInsightsId = array_search('INFO', array_column($SOCInsights->result->data, 'InsightsAggregated.priorityText'));
				$LowInsightsId = array_search('LOW', array_column($SOCInsights->result->data, 'InsightsAggregated.priorityText'));
				$MediumInsightsId = array_search('MEDIUM', array_column($SOCInsights->result->data, 'InsightsAggregated.priorityText'));
				$HighInsightsId = array_search('HIGH', array_column($SOCInsights->result->data, 'InsightsAggregated.priorityText'));
				$CriticalInsightsId = array_search('CRITICAL', array_column($SOCInsights->result->data, 'InsightsAggregated.priorityText'));
				$TotalInsights = number_abbr(array_sum(array_column($SOCInsights->result->data, 'InsightsAggregated.count')));
			} else {
				$TotalInsights = 0;
			}
			if (isset($InfoInsightsId) AND $InfoInsightsId !== false) {$InfoInsights = $SOCInsights->result->data[$InfoInsightsId]->{'InsightsAggregated.count'};} else {$InfoInsights = 0;}
			if (isset($LowInsightsId) AND $LowInsightsId !== false) {$LowInsights = $SOCInsights->result->data[$LowInsightsId]->{'InsightsAggregated.count'};} else {$LowInsights = 0;}
			if (isset($MediumInsightsId) AND $MediumInsightsId !== false) {$MediumInsights = $SOCInsights->result->data[$MediumInsightsId]->{'InsightsAggregated.count'};} else {$MediumInsights = 0;}
			if (isset($HighInsightsId) AND $HighInsightsId !== false) {$HighInsights = $SOCInsights->result->data[$HighInsightsId]->{'InsightsAggregated.count'};} else {$HighInsights = 0;}
			if (isset($CriticalInsightsId) AND $CriticalInsightsId !== false) {$CriticalInsights = $SOCInsights->result->data[$CriticalInsightsId]->{'InsightsAggregated.count'};} else {$CriticalInsights = 0;}
	
			// Security Activity
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Security Activity");
			$SecurityEvents = $CubeJSResults['SecurityEvents']['Body'];
			if (isset($SecurityEvents->result->data[0])) {
				$SecurityEventsCount = $SecurityEvents->result->data[0]->{'PortunusAggInsight.requests'};
			} else {
				$SecurityEventsCount = 0;
			}
	
			// Data Exfiltration Events
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Data Exfiltration Events");
			$DataExfilEvents = $CubeJSResults['DataExfilEvents']['Body'];
			if (isset($DataExfilEvents->result->data[0])) {
				$DataExfilEventsCount = $DataExfilEvents->result->data[0]->{'PortunusAggInsight.requests'};
			} else {
				$DataExfilEventsCount = 0;
			}
	
			// Zero Day DNS Events
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Zero Day DNS Events");
			$ZeroDayDNSEvents = $CubeJSResults['ZeroDayDNSEvents']['Body'];
			if (isset($ZeroDayDNSEvents->result->data[0])) {
				$ZeroDayDNSEventsCount = $ZeroDayDNSEvents->result->data[0]->{'PortunusAggInsight.requests'};
			} else {
				$ZeroDayDNSEventsCount = 0;
			}
	
			// Suspicious Domains
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Suspicious Domain Events");
			$SuspiciousEvents = $CubeJSResults['SuspiciousEvents']['Body'];
			if (isset($SuspiciousEvents->result->data[0])) {
				$SuspiciousEventsCount = $SuspiciousEvents->result->data[0]->{'PortunusAggInsight.requests'};
			} else {
				$SuspiciousEventsCount = 0;
			}
	
			// High Risk Websites
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building High Risk Website Events");
			$HighRiskWebsites = $CubeJSResults['HighRiskWebsites']['Body'];
			if (isset($HighRiskWebsites->result->data)) {
				$HighRiskWebsiteCount = array_sum(array_column($HighRiskWebsites->result->data, 'PortunusAggWebContentDiscovery.count'));
				$HighRiskWebCategoryCount = count($HighRiskWebsites->result->data);
			} else {
				$HighRiskWebsiteCount = 0;
				$HighRiskWebCategoryCount = 0;
			}
	
			// DNS over HTTPS
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building DoH Events");
			$DOHEvents = $CubeJSResults['DOHEvents']['Body'];
			if (isset($DOHEvents->result->data[0])) {
				$DOHEventsCount = $DOHEvents->result->data[0]->{'PortunusAggInsight.requests'};
			} else {
				$DOHEventsCount = 0;
			}
	
			// Newly Observed Domains
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Newly Observed Domain Events");
			$NODEvents = $CubeJSResults['NODEvents']['Body'];
			if (isset($NODEvents->result->data[0])) {
				$NODEventsCount = $NODEvents->result->data[0]->{'PortunusAggInsight.requests'};
			} else {
				$NODEventsCount = 0;
			}
	
			// Domain Generation Algorithms
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building DGA Events");
			$DGAEvents = $CubeJSResults['DGAEvents']['Body'];
			if (isset($DGAEvents->result->data[0])) {
				$DGAEventsCount = $DGAEvents->result->data[0]->{'PortunusAggInsight.requests'};
			} else {
				$DGAEventsCount = 0;
			}
	
			// App Discovery - Unique Applications
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building list of Unique Applications");
			$AppDiscoveryApplications = $CubeJSResults['AppDiscoveryApplications']['Body'];
			if (isset($AppDiscoveryApplications->result->data)) {
				$AppDiscoveryApplicationsCount = count($AppDiscoveryApplications->result->data);
			} else {
				$AppDiscoveryApplicationsCount = 0;
			}

			// App Discovery - Total Requests / Devices
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Collecting totals from application discovery");
			$AppDiscoveryTotals = $CubeJSResults['AppDiscoveryTotals']['Body'];
			if (isset($AppDiscoveryTotals->result->data[0])) {
				$AppDiscoveryTotalRequestsCount = $AppDiscoveryTotals->result->data[0]->{'PortunusAggAppDiscovery.requests'};
				$AppDiscoveryTotalDevicesCount = $AppDiscoveryTotals->result->data[0]->{'PortunusAggAppDiscovery.deviceCount'};
			} else {
				$AppDiscoveryTotalRequestsCount = 0;
				$AppDiscoveryTotalDevicesCount = 0;
			}

			// Web Content - Total Requests / Devices
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Web Content Request Count");
			$WebContentTotals = $CubeJSResults['WebContentTotals']['Body'];
			if (isset($WebContentTotals->result->data[0])) {
				$WebContentTotalRequestsCount = $WebContentTotals->result->data[0]->{'PortunusAggWebContentDiscovery.requests'};
				$WebContentTotalDevicesCount = $WebContentTotals->result->data[0]->{'PortunusAggWebContentDiscovery.deviceCount'};
			} else {
				$WebContentTotalRequestsCount = 0;
				$WebContentTotalDevicesCount = 0;
			}

			// Web Content - Total Security Events
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Web Content Security Event Count");
			$WebContentSecurityEvents = $CubeJSResults['WebContentSecurityEvents']['Body'];
			if (isset($WebContentSecurityEvents->result->data[0])) {
				$WebContentSecurityEventsCount = $WebContentSecurityEvents->result->data[0]->{'PortunusAggWebcontent.requests'};
			} else {
				$WebContentSecurityEventsCount = 0;
			}

			// Web Content - Content Categories
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Web Content Categories");
			$WebContentCategoriesResult = $this->queryCSP("get","api/atcfw/v1/content_categories");
			$WebContentCategories = [];
			if (isset($WebContentCategoriesResult->results) && is_array($WebContentCategoriesResult->results)) {
				foreach ($WebContentCategoriesResult->results as $Category) {
					$WebContentCategories[$Category->category_name] = [
						'Code' => $Category->category_code,
						'Group' => $Category->functional_group
					];
				}
			}

			// Example response
			// {
			// 	"results": [
			// 		{
			// 			"category_code": 10001,
			// 			"category_name": "Abortion",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10002,
			// 			"category_name": "Abortion Pro Choice",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10003,
			// 			"category_name": "Abortion Pro Life",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10004,
			// 			"category_name": "Child Inappropriate",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10005,
			// 			"category_name": "Gambling",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10006,
			// 			"category_name": "Gay, Lesbian or Bisexual",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10007,
			// 			"category_name": "Lingerie, Suggestive \u0026 Pinup",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10008,
			// 			"category_name": "Nudity",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10009,
			// 			"category_name": "Pornography",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10010,
			// 			"category_name": "Profanity",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10011,
			// 			"category_name": "R-Rated",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10012,
			// 			"category_name": "Sex \u0026 Erotic",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10013,
			// 			"category_name": "Sex Education",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10014,
			// 			"category_name": "Tobacco",
			// 			"functional_group": "Adult"
			// 		},
			// 		{
			// 			"category_code": 10015,
			// 			"category_name": "Military",
			// 			"functional_group": "Aggressive"
			// 		},
			// 		{
			// 			"category_code": 10016,
			// 			"category_name": "Violence",
			// 			"functional_group": "Aggressive"
			// 		},
			// 		{
			// 			"category_code": 10017,
			// 			"category_name": "Weapons",
			// 			"functional_group": "Aggressive"
			// 		},
			// 		{
			// 			"category_code": 10018,
			// 			"category_name": "Aggressive - Other",
			// 			"functional_group": "Aggressive"
			// 		},
			// 		{
			// 			"category_code": 10019,
			// 			"category_name": "Fine Art",
			// 			"functional_group": "Arts"
			// 		},
			// 		{
			// 			"category_code": 10020,
			// 			"category_name": "Arts - Other",
			// 			"functional_group": "Arts"
			// 		},
			// 		{
			// 			"category_code": 10021,
			// 			"category_name": "Auto Parts",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10022,
			// 			"category_name": "Auto Repair",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10023,
			// 			"category_name": "Buying/Selling Cars",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10024,
			// 			"category_name": "Car Culture",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10025,
			// 			"category_name": "Certified Pre-Owned",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10026,
			// 			"category_name": "Convertible",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10027,
			// 			"category_name": "Coupe",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10028,
			// 			"category_name": "Crossover",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10029,
			// 			"category_name": "Diesel",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10030,
			// 			"category_name": "Electric Vehicle",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10031,
			// 			"category_name": "Hatchback",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10032,
			// 			"category_name": "Hybrid",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10033,
			// 			"category_name": "Luxury",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10034,
			// 			"category_name": "MiniVan",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10035,
			// 			"category_name": "Motorcycles",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10036,
			// 			"category_name": "Off-Road Vehicles",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10037,
			// 			"category_name": "Performance Vehicles",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10038,
			// 			"category_name": "Pickup",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10039,
			// 			"category_name": "Road-Side Assistance",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10040,
			// 			"category_name": "Sedan",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10041,
			// 			"category_name": "Trucks \u0026 Accessories",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10042,
			// 			"category_name": "Vintage Cars",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10043,
			// 			"category_name": "Wagon",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10044,
			// 			"category_name": "Automotive - Other",
			// 			"functional_group": "Automotive"
			// 		},
			// 		{
			// 			"category_code": 10045,
			// 			"category_name": "Agriculture",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10046,
			// 			"category_name": "Biotechnology",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10047,
			// 			"category_name": "Business Software",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10048,
			// 			"category_name": "Construction",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10049,
			// 			"category_name": "Forestry",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10050,
			// 			"category_name": "Government",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10051,
			// 			"category_name": "Green Solutions \u0026 Conservation",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10052,
			// 			"category_name": "Home \u0026 Office Furnishings",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10053,
			// 			"category_name": "Human Resources",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10054,
			// 			"category_name": "Manufacturing",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10055,
			// 			"category_name": "Marketing Services",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10056,
			// 			"category_name": "Metals",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10057,
			// 			"category_name": "Physical Security",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10058,
			// 			"category_name": "Productivity",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10059,
			// 			"category_name": "Retirement Homes \u0026 Assisted Living",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10060,
			// 			"category_name": "Shipping \u0026 Logistics",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10061,
			// 			"category_name": "Business - Other",
			// 			"functional_group": "Business"
			// 		},
			// 		{
			// 			"category_code": 10062,
			// 			"category_name": "Career Advice",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10063,
			// 			"category_name": "Career Planning",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10064,
			// 			"category_name": "College",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10065,
			// 			"category_name": "Financial Aid",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10066,
			// 			"category_name": "Job Fairs",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10067,
			// 			"category_name": "Job Search",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10068,
			// 			"category_name": "Nursing",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10069,
			// 			"category_name": "Resume Writing/Advice",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10070,
			// 			"category_name": "Scholarships",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10071,
			// 			"category_name": "Telecommuting",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10072,
			// 			"category_name": "U.S. Military",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10073,
			// 			"category_name": "Careers - Other",
			// 			"functional_group": "Careers"
			// 		},
			// 		{
			// 			"category_code": 10074,
			// 			"category_name": "Child Abuse Images",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10075,
			// 			"category_name": "Criminal Skills",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10490,
			// 			"category_name": "Terrorism",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10076,
			// 			"category_name": "Hacking",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10077,
			// 			"category_name": "Hate Speech",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10078,
			// 			"category_name": "Illegal Drugs",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10079,
			// 			"category_name": "Marijuana",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10080,
			// 			"category_name": "Piracy \u0026 Copyright Theft",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10081,
			// 			"category_name": "School Cheating",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10082,
			// 			"category_name": "Self Harm",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10083,
			// 			"category_name": "Torrent Repository",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10084,
			// 			"category_name": "Criminal Activities - Other",
			// 			"functional_group": "Criminal Activities"
			// 		},
			// 		{
			// 			"category_code": 10085,
			// 			"category_name": "Anonymizer",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10086,
			// 			"category_name": "Chat",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10087,
			// 			"category_name": "Community Forums",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10088,
			// 			"category_name": "Instant Messenger",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10089,
			// 			"category_name": "Login Screens",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10090,
			// 			"category_name": "Personal Pages \u0026 Blogs",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10091,
			// 			"category_name": "Photo Sharing",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10092,
			// 			"category_name": "Professional Networking",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10093,
			// 			"category_name": "Redirect",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10094,
			// 			"category_name": "Social Networking",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10095,
			// 			"category_name": "Text Messaging \u0026 SMS",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10096,
			// 			"category_name": "Translator",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10097,
			// 			"category_name": "Web-based Email",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10098,
			// 			"category_name": "Web-based Greeting Cards",
			// 			"functional_group": "Dynamic"
			// 		},
			// 		{
			// 			"category_code": 10099,
			// 			"category_name": "7-12 Education",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10100,
			// 			"category_name": "Adult Education",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10101,
			// 			"category_name": "Art History",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10102,
			// 			"category_name": "College Administration",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10103,
			// 			"category_name": "College Life",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10104,
			// 			"category_name": "Distance Learning",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10105,
			// 			"category_name": "Educational Institutions",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10106,
			// 			"category_name": "Educational Materials \u0026 Studies",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10107,
			// 			"category_name": "English as a 2nd Language",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10108,
			// 			"category_name": "Graduate School",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10109,
			// 			"category_name": "Homeschooling",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10110,
			// 			"category_name": "Homework/Study Tips",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10111,
			// 			"category_name": "K-6 Educators",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10112,
			// 			"category_name": "Language Learning",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10113,
			// 			"category_name": "Literature \u0026 Books",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10114,
			// 			"category_name": "Private School",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10115,
			// 			"category_name": "Reference Materials \u0026 Maps",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10116,
			// 			"category_name": "Special Education",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10117,
			// 			"category_name": "Studying Business",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10118,
			// 			"category_name": "Tutoring",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10119,
			// 			"category_name": "Wikis",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10120,
			// 			"category_name": "Education - Other",
			// 			"functional_group": "Education"
			// 		},
			// 		{
			// 			"category_code": 10121,
			// 			"category_name": "Entertainment News \u0026 Celebrity Sites",
			// 			"functional_group": "Entertainment"
			// 		},
			// 		{
			// 			"category_code": 10122,
			// 			"category_name": "Entertainment Venues \u0026 Events",
			// 			"functional_group": "Entertainment"
			// 		},
			// 		{
			// 			"category_code": 10123,
			// 			"category_name": "Humor",
			// 			"functional_group": "Entertainment"
			// 		},
			// 		{
			// 			"category_code": 10124,
			// 			"category_name": "Movies",
			// 			"functional_group": "Entertainment"
			// 		},
			// 		{
			// 			"category_code": 10125,
			// 			"category_name": "Music",
			// 			"functional_group": "Entertainment"
			// 		},
			// 		{
			// 			"category_code": 10126,
			// 			"category_name": "Streaming \u0026 Downloadable Audio",
			// 			"functional_group": "Entertainment"
			// 		},
			// 		{
			// 			"category_code": 10127,
			// 			"category_name": "Streaming \u0026 Downloadable Video",
			// 			"functional_group": "Entertainment"
			// 		},
			// 		{
			// 			"category_code": 10128,
			// 			"category_name": "Television",
			// 			"functional_group": "Entertainment"
			// 		},
			// 		{
			// 			"category_code": 10129,
			// 			"category_name": "Entertainment - Other",
			// 			"functional_group": "Entertainment"
			// 		},
			// 		{
			// 			"category_code": 10130,
			// 			"category_name": "Adoption",
			// 			"functional_group": "Family \u0026 Parenting"
			// 		},
			// 		{
			// 			"category_code": 10131,
			// 			"category_name": "Babies and Toddlers",
			// 			"functional_group": "Family \u0026 Parenting"
			// 		},
			// 		{
			// 			"category_code": 10132,
			// 			"category_name": "Daycare/Pre School",
			// 			"functional_group": "Family \u0026 Parenting"
			// 		},
			// 		{
			// 			"category_code": 10133,
			// 			"category_name": "Eldercare",
			// 			"functional_group": "Family \u0026 Parenting"
			// 		},
			// 		{
			// 			"category_code": 10134,
			// 			"category_name": "Family Internet",
			// 			"functional_group": "Family \u0026 Parenting"
			// 		},
			// 		{
			// 			"category_code": 10135,
			// 			"category_name": "Parenting - K-6 Kids",
			// 			"functional_group": "Family \u0026 Parenting"
			// 		},
			// 		{
			// 			"category_code": 10136,
			// 			"category_name": "Parenting Teens",
			// 			"functional_group": "Family \u0026 Parenting"
			// 		},
			// 		{
			// 			"category_code": 10137,
			// 			"category_name": "Pregnancy",
			// 			"functional_group": "Family \u0026 Parenting"
			// 		},
			// 		{
			// 			"category_code": 10138,
			// 			"category_name": "Special Needs Kids",
			// 			"functional_group": "Family \u0026 Parenting"
			// 		},
			// 		{
			// 			"category_code": 10139,
			// 			"category_name": "Family \u0026 Parenting - Other",
			// 			"functional_group": "Family \u0026 Parenting"
			// 		},
			// 		{
			// 			"category_code": 10140,
			// 			"category_name": "Accessories",
			// 			"functional_group": "Fashion"
			// 		},
			// 		{
			// 			"category_code": 10141,
			// 			"category_name": "Beauty",
			// 			"functional_group": "Fashion"
			// 		},
			// 		{
			// 			"category_code": 10142,
			// 			"category_name": "Body Art",
			// 			"functional_group": "Fashion"
			// 		},
			// 		{
			// 			"category_code": 10143,
			// 			"category_name": "Clothing",
			// 			"functional_group": "Fashion"
			// 		},
			// 		{
			// 			"category_code": 10144,
			// 			"category_name": "Fashion",
			// 			"functional_group": "Fashion"
			// 		},
			// 		{
			// 			"category_code": 10145,
			// 			"category_name": "Jewelry",
			// 			"functional_group": "Fashion"
			// 		},
			// 		{
			// 			"category_code": 10146,
			// 			"category_name": "Swimsuits",
			// 			"functional_group": "Fashion"
			// 		},
			// 		{
			// 			"category_code": 10147,
			// 			"category_name": "Fashion - Other",
			// 			"functional_group": "Fashion"
			// 		},
			// 		{
			// 			"category_code": 10148,
			// 			"category_name": "Accounting",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10149,
			// 			"category_name": "Banking",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10150,
			// 			"category_name": "Beginning Investing",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10151,
			// 			"category_name": "Credit/Debt \u0026 Loans",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10152,
			// 			"category_name": "Financial News",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10153,
			// 			"category_name": "Financial Planning",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10154,
			// 			"category_name": "Hedge Fund",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10155,
			// 			"category_name": "Insurance",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10156,
			// 			"category_name": "Investing",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10157,
			// 			"category_name": "Mutual Funds",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10158,
			// 			"category_name": "Online Financial Tools \u0026 Quotes",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10159,
			// 			"category_name": "Options",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10160,
			// 			"category_name": "Retirement Planning",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10161,
			// 			"category_name": "Stocks",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10162,
			// 			"category_name": "Tax Planning",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10163,
			// 			"category_name": "Finance - Other",
			// 			"functional_group": "Finance"
			// 		},
			// 		{
			// 			"category_code": 10164,
			// 			"category_name": "American Cuisine",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10165,
			// 			"category_name": "Barbecues \u0026 Grilling",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10166,
			// 			"category_name": "Cajun/Creole",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10167,
			// 			"category_name": "Chinese Cuisine",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10168,
			// 			"category_name": "Cocktails/Beer",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10169,
			// 			"category_name": "Coffee/Tea",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10170,
			// 			"category_name": "Cuisine-Specific",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10171,
			// 			"category_name": "Desserts \u0026 Baking",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10172,
			// 			"category_name": "Dining Out",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10173,
			// 			"category_name": "Food Allergies",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10174,
			// 			"category_name": "French Cuisine",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10175,
			// 			"category_name": "Health/Lowfat Cooking",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10176,
			// 			"category_name": "Italian Cuisine",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10177,
			// 			"category_name": "Japanese Cuisine",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10178,
			// 			"category_name": "Mexican Cuisine",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10179,
			// 			"category_name": "Vegan",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10180,
			// 			"category_name": "Vegetarian",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10181,
			// 			"category_name": "Wine",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10182,
			// 			"category_name": "Food \u0026 Drink - Other",
			// 			"functional_group": "Food \u0026 Drink"
			// 		},
			// 		{
			// 			"category_code": 10183,
			// 			"category_name": "A.D.D.",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10184,
			// 			"category_name": "AIDS/HIV",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10185,
			// 			"category_name": "Allergies",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10186,
			// 			"category_name": "Alternative Medicine",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10187,
			// 			"category_name": "Arthritis",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10188,
			// 			"category_name": "Asthma",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10189,
			// 			"category_name": "Autism/PDD",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10190,
			// 			"category_name": "Bipolar Disorder",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10191,
			// 			"category_name": "Brain Tumor",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10192,
			// 			"category_name": "Cancer",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10193,
			// 			"category_name": "Children's Health",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10194,
			// 			"category_name": "Cholesterol",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10195,
			// 			"category_name": "Chronic Fatigue",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10196,
			// 			"category_name": "Chronic Pain",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10197,
			// 			"category_name": "Cold \u0026 Flu",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10198,
			// 			"category_name": "Cosmetic Surgery",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10199,
			// 			"category_name": "Deafness",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10200,
			// 			"category_name": "Dental Care",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10201,
			// 			"category_name": "Depression",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10202,
			// 			"category_name": "Dermatology",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10203,
			// 			"category_name": "Diabetes",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10204,
			// 			"category_name": "Disorders",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10205,
			// 			"category_name": "Epilepsy",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10206,
			// 			"category_name": "Exercise",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10207,
			// 			"category_name": "GERD/Acid Reflux",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10208,
			// 			"category_name": "Headaches/Migraines",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10209,
			// 			"category_name": "Heart Disease",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10210,
			// 			"category_name": "Herbs for Health",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10211,
			// 			"category_name": "Holistic Healing",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10212,
			// 			"category_name": "IBS/Crohn's Disease",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10213,
			// 			"category_name": "Incest/Abuse Support",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10214,
			// 			"category_name": "Incontinence",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10215,
			// 			"category_name": "Infertility",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10216,
			// 			"category_name": "Men's Health",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10217,
			// 			"category_name": "Nutrition \u0026 Diet",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10218,
			// 			"category_name": "Orthopedics",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10219,
			// 			"category_name": "Panic/Anxiety",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10220,
			// 			"category_name": "Pediatrics",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10221,
			// 			"category_name": "Pharmaceuticals",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10222,
			// 			"category_name": "Physical Therapy",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10223,
			// 			"category_name": "Psychology/Psychiatry",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10224,
			// 			"category_name": "Self-help \u0026 Addiction",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10225,
			// 			"category_name": "Senior Health",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10226,
			// 			"category_name": "Sexuality",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10227,
			// 			"category_name": "Sleep Disorders",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10228,
			// 			"category_name": "Smoking Cessation",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10229,
			// 			"category_name": "Supplements \u0026 Compounds",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10230,
			// 			"category_name": "Syndrome",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10231,
			// 			"category_name": "Thyroid Disease",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10232,
			// 			"category_name": "Weight Loss",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10233,
			// 			"category_name": "Women's Health",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10234,
			// 			"category_name": "Health - Other",
			// 			"functional_group": "Health"
			// 		},
			// 		{
			// 			"category_code": 10235,
			// 			"category_name": "Art/Technology",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10236,
			// 			"category_name": "Arts \u0026 Crafts",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10237,
			// 			"category_name": "Beadwork",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10238,
			// 			"category_name": "Birdwatching",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10239,
			// 			"category_name": "Board Games/Puzzles",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10240,
			// 			"category_name": "Candle \u0026 Soap Making",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10241,
			// 			"category_name": "Card Games",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10242,
			// 			"category_name": "Cartoons \u0026 Anime",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10243,
			// 			"category_name": "Chess",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10244,
			// 			"category_name": "Cigars",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10245,
			// 			"category_name": "Collecting",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10246,
			// 			"category_name": "Comic Books",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10247,
			// 			"category_name": "Drawing/Sketching",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10248,
			// 			"category_name": "Freelance Writing",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10249,
			// 			"category_name": "Genealogy",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10250,
			// 			"category_name": "Getting Published",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10251,
			// 			"category_name": "Guitar",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10252,
			// 			"category_name": "Home Recording",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10253,
			// 			"category_name": "Investors \u0026 Patents",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10254,
			// 			"category_name": "Jewelry Making",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10255,
			// 			"category_name": "Magic \u0026 Illusion",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10256,
			// 			"category_name": "Needlework",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10257,
			// 			"category_name": "Painting",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10258,
			// 			"category_name": "Photography",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10259,
			// 			"category_name": "Radio",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10260,
			// 			"category_name": "Roleplaying Games",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10261,
			// 			"category_name": "Sci-Fi \u0026 Fantasy",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10262,
			// 			"category_name": "Scrapbooking",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10263,
			// 			"category_name": "Screenwriting",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10264,
			// 			"category_name": "Stamps \u0026 Coins",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10265,
			// 			"category_name": "Themes",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10266,
			// 			"category_name": "Video \u0026 Computer Games",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10267,
			// 			"category_name": "Woodworking",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10268,
			// 			"category_name": "Hobbies \u0026 Interests - Other",
			// 			"functional_group": "Hobbies \u0026 Interests"
			// 		},
			// 		{
			// 			"category_code": 10269,
			// 			"category_name": "Appliances",
			// 			"functional_group": "Home \u0026 Garden"
			// 		},
			// 		{
			// 			"category_code": 10270,
			// 			"category_name": "Entertaining",
			// 			"functional_group": "Home \u0026 Garden"
			// 		},
			// 		{
			// 			"category_code": 10271,
			// 			"category_name": "Environmental Safety",
			// 			"functional_group": "Home \u0026 Garden"
			// 		},
			// 		{
			// 			"category_code": 10272,
			// 			"category_name": "Gardening",
			// 			"functional_group": "Home \u0026 Garden"
			// 		},
			// 		{
			// 			"category_code": 10273,
			// 			"category_name": "Home Repair",
			// 			"functional_group": "Home \u0026 Garden"
			// 		},
			// 		{
			// 			"category_code": 10274,
			// 			"category_name": "Home Theater",
			// 			"functional_group": "Home \u0026 Garden"
			// 		},
			// 		{
			// 			"category_code": 10275,
			// 			"category_name": "Interior Decorating",
			// 			"functional_group": "Home \u0026 Garden"
			// 		},
			// 		{
			// 			"category_code": 10276,
			// 			"category_name": "Landscaping",
			// 			"functional_group": "Home \u0026 Garden"
			// 		},
			// 		{
			// 			"category_code": 10277,
			// 			"category_name": "Remodeling \u0026 Construction",
			// 			"functional_group": "Home \u0026 Garden"
			// 		},
			// 		{
			// 			"category_code": 10278,
			// 			"category_name": "Home \u0026 Garden - Other",
			// 			"functional_group": "Home \u0026 Garden"
			// 		},
			// 		{
			// 			"category_code": 10279,
			// 			"category_name": "Games",
			// 			"functional_group": "Kids"
			// 		},
			// 		{
			// 			"category_code": 10280,
			// 			"category_name": "Kid's Pages",
			// 			"functional_group": "Kids"
			// 		},
			// 		{
			// 			"category_code": 10281,
			// 			"category_name": "Toys",
			// 			"functional_group": "Kids"
			// 		},
			// 		{
			// 			"category_code": 10282,
			// 			"category_name": "Kids - Other",
			// 			"functional_group": "Kids"
			// 		},
			// 		{
			// 			"category_code": 10283,
			// 			"category_name": "Dating \u0026 Relationships",
			// 			"functional_group": "Lifestyle"
			// 		},
			// 		{
			// 			"category_code": 10284,
			// 			"category_name": "Divorce Support",
			// 			"functional_group": "Lifestyle"
			// 		},
			// 		{
			// 			"category_code": 10285,
			// 			"category_name": "Ethnic Specific",
			// 			"functional_group": "Lifestyle"
			// 		},
			// 		{
			// 			"category_code": 10286,
			// 			"category_name": "Marriage",
			// 			"functional_group": "Lifestyle"
			// 		},
			// 		{
			// 			"category_code": 10287,
			// 			"category_name": "Parks, Rec Facilities \u0026 Gyms",
			// 			"functional_group": "Lifestyle"
			// 		},
			// 		{
			// 			"category_code": 10288,
			// 			"category_name": "Senior Living",
			// 			"functional_group": "Lifestyle"
			// 		},
			// 		{
			// 			"category_code": 10289,
			// 			"category_name": "Teens",
			// 			"functional_group": "Lifestyle"
			// 		},
			// 		{
			// 			"category_code": 10290,
			// 			"category_name": "Weddings",
			// 			"functional_group": "Lifestyle"
			// 		},
			// 		{
			// 			"category_code": 10291,
			// 			"category_name": "Lifestyle - Other",
			// 			"functional_group": "Lifestyle"
			// 		},
			// 		{
			// 			"category_code": 10292,
			// 			"category_name": "Ad Fraud",
			// 			"functional_group": "Malicious"
			// 		},
			// 		{
			// 			"category_code": 10293,
			// 			"category_name": "Botnet",
			// 			"functional_group": "Malicious"
			// 		},
			// 		{
			// 			"category_code": 10294,
			// 			"category_name": "Command and Control Centers",
			// 			"functional_group": "Malicious"
			// 		},
			// 		{
			// 			"category_code": 10295,
			// 			"category_name": "Compromised \u0026 Links To Malware",
			// 			"functional_group": "Malicious"
			// 		},
			// 		{
			// 			"category_code": 10296,
			// 			"category_name": "Malware Call-Home",
			// 			"functional_group": "Malicious"
			// 		},
			// 		{
			// 			"category_code": 10297,
			// 			"category_name": "Malware Distribution Point",
			// 			"functional_group": "Malicious"
			// 		},
			// 		{
			// 			"category_code": 10298,
			// 			"category_name": "Phishing/Fraud",
			// 			"functional_group": "Malicious"
			// 		},
			// 		{
			// 			"category_code": 10299,
			// 			"category_name": "Spam URLs",
			// 			"functional_group": "Malicious"
			// 		},
			// 		{
			// 			"category_code": 10492,
			// 			"category_name": "Cryptocurrency Mining",
			// 			"functional_group": "Malicious"
			// 		},
			// 		{
			// 			"category_code": 10300,
			// 			"category_name": "Spyware \u0026 Questionable Software",
			// 			"functional_group": "Malicious"
			// 		},
			// 		{
			// 			"category_code": 10301,
			// 			"category_name": "Content Server",
			// 			"functional_group": "Miscellaneous"
			// 		},
			// 		{
			// 			"category_code": 10302,
			// 			"category_name": "No Content Found",
			// 			"functional_group": "Miscellaneous"
			// 		},
			// 		{
			// 			"category_code": 10303,
			// 			"category_name": "Parked \u0026 For Sale Domains",
			// 			"functional_group": "Miscellaneous"
			// 		},
			// 		{
			// 			"category_code": 10304,
			// 			"category_name": "Private IP Address",
			// 			"functional_group": "Miscellaneous"
			// 		},
			// 		{
			// 			"category_code": 10494,
			// 			"category_name": "Fake News",
			// 			"functional_group": "Miscellaneous"
			// 		},
			// 		{
			// 			"category_code": 10305,
			// 			"category_name": "Unreachable",
			// 			"functional_group": "Miscellaneous"
			// 		},
			// 		{
			// 			"category_code": 10306,
			// 			"category_name": "Miscellaneous - Other",
			// 			"functional_group": "Miscellaneous"
			// 		},
			// 		{
			// 			"category_code": 10307,
			// 			"category_name": "Image Search",
			// 			"functional_group": "News, Portal \u0026 Search"
			// 		},
			// 		{
			// 			"category_code": 10308,
			// 			"category_name": "International News",
			// 			"functional_group": "News, Portal \u0026 Search"
			// 		},
			// 		{
			// 			"category_code": 10309,
			// 			"category_name": "Local News",
			// 			"functional_group": "News, Portal \u0026 Search"
			// 		},
			// 		{
			// 			"category_code": 10310,
			// 			"category_name": "Magazines",
			// 			"functional_group": "News, Portal \u0026 Search"
			// 		},
			// 		{
			// 			"category_code": 10311,
			// 			"category_name": "National News",
			// 			"functional_group": "News, Portal \u0026 Search"
			// 		},
			// 		{
			// 			"category_code": 10312,
			// 			"category_name": "Portal Sites",
			// 			"functional_group": "News, Portal \u0026 Search"
			// 		},
			// 		{
			// 			"category_code": 10313,
			// 			"category_name": "Search Engines",
			// 			"functional_group": "News, Portal \u0026 Search"
			// 		},
			// 		{
			// 			"category_code": 10314,
			// 			"category_name": "News, Portal \u0026 Search - Other",
			// 			"functional_group": "News, Portal \u0026 Search"
			// 		},
			// 		{
			// 			"category_code": 10315,
			// 			"category_name": "Pay To Surf",
			// 			"functional_group": "Online Ads"
			// 		},
			// 		{
			// 			"category_code": 10316,
			// 			"category_name": "Online Ads - Other",
			// 			"functional_group": "Online Ads"
			// 		},
			// 		{
			// 			"category_code": 10317,
			// 			"category_name": "Aquariums",
			// 			"functional_group": "Pets"
			// 		},
			// 		{
			// 			"category_code": 10318,
			// 			"category_name": "Birds",
			// 			"functional_group": "Pets"
			// 		},
			// 		{
			// 			"category_code": 10319,
			// 			"category_name": "Cats",
			// 			"functional_group": "Pets"
			// 		},
			// 		{
			// 			"category_code": 10320,
			// 			"category_name": "Dogs",
			// 			"functional_group": "Pets"
			// 		},
			// 		{
			// 			"category_code": 10321,
			// 			"category_name": "Large Animals",
			// 			"functional_group": "Pets"
			// 		},
			// 		{
			// 			"category_code": 10322,
			// 			"category_name": "Reptiles",
			// 			"functional_group": "Pets"
			// 		},
			// 		{
			// 			"category_code": 10323,
			// 			"category_name": "Veterinary Medicine",
			// 			"functional_group": "Pets"
			// 		},
			// 		{
			// 			"category_code": 10324,
			// 			"category_name": "Pets - Other",
			// 			"functional_group": "Pets"
			// 		},
			// 		{
			// 			"category_code": 10325,
			// 			"category_name": "Advocacy Groups \u0026 Trade Associations",
			// 			"functional_group": "Public, Government \u0026 Law"
			// 		},
			// 		{
			// 			"category_code": 10326,
			// 			"category_name": "Commentary",
			// 			"functional_group": "Public, Government \u0026 Law"
			// 		},
			// 		{
			// 			"category_code": 10327,
			// 			"category_name": "Government Sponsored",
			// 			"functional_group": "Public, Government \u0026 Law"
			// 		},
			// 		{
			// 			"category_code": 10328,
			// 			"category_name": "Immigration",
			// 			"functional_group": "Public, Government \u0026 Law"
			// 		},
			// 		{
			// 			"category_code": 10329,
			// 			"category_name": "Legal Issues",
			// 			"functional_group": "Public, Government \u0026 Law"
			// 		},
			// 		{
			// 			"category_code": 10330,
			// 			"category_name": "Philanthropic Organizations",
			// 			"functional_group": "Public, Government \u0026 Law"
			// 		},
			// 		{
			// 			"category_code": 10331,
			// 			"category_name": "Politics",
			// 			"functional_group": "Public, Government \u0026 Law"
			// 		},
			// 		{
			// 			"category_code": 10332,
			// 			"category_name": "Social \u0026 Affiliation Organizations",
			// 			"functional_group": "Public, Government \u0026 Law"
			// 		},
			// 		{
			// 			"category_code": 10333,
			// 			"category_name": "U.S. Government Resources",
			// 			"functional_group": "Public, Government \u0026 Law"
			// 		},
			// 		{
			// 			"category_code": 10334,
			// 			"category_name": "Public, Government \u0026 Law - Other",
			// 			"functional_group": "Public, Government \u0026 Law"
			// 		},
			// 		{
			// 			"category_code": 10335,
			// 			"category_name": "Apartments",
			// 			"functional_group": "Real Estate"
			// 		},
			// 		{
			// 			"category_code": 10336,
			// 			"category_name": "Architects",
			// 			"functional_group": "Real Estate"
			// 		},
			// 		{
			// 			"category_code": 10337,
			// 			"category_name": "Buying/Selling Homes",
			// 			"functional_group": "Real Estate"
			// 		},
			// 		{
			// 			"category_code": 10338,
			// 			"category_name": "Real Estate - Other",
			// 			"functional_group": "Real Estate"
			// 		},
			// 		{
			// 			"category_code": 10339,
			// 			"category_name": "Alternative Religions",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10340,
			// 			"category_name": "Atheism \u0026 Agnosticism",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10341,
			// 			"category_name": "Buddhism",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10342,
			// 			"category_name": "Catholicism",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10343,
			// 			"category_name": "Christianity",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10344,
			// 			"category_name": "Hinduism",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10345,
			// 			"category_name": "Islam",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10346,
			// 			"category_name": "Judaism",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10347,
			// 			"category_name": "Latter-Day Saints",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10348,
			// 			"category_name": "Occult",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10349,
			// 			"category_name": "Pagan/Wiccan",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10350,
			// 			"category_name": "Religion - Other",
			// 			"functional_group": "Religion"
			// 		},
			// 		{
			// 			"category_code": 10351,
			// 			"category_name": "Anatomy",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10352,
			// 			"category_name": "Astrology \u0026 Horoscopes",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10353,
			// 			"category_name": "Biology",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10354,
			// 			"category_name": "Botany",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10355,
			// 			"category_name": "Chemistry",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10361,
			// 			"category_name": "Weather",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10356,
			// 			"category_name": "Geography",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10357,
			// 			"category_name": "Geology",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10358,
			// 			"category_name": "Paranormal Phenomena",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10359,
			// 			"category_name": "Physics",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10360,
			// 			"category_name": "Space/Astronomy",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10362,
			// 			"category_name": "Science - Other",
			// 			"functional_group": "Science"
			// 		},
			// 		{
			// 			"category_code": 10363,
			// 			"category_name": "Auctions \u0026 Marketplaces",
			// 			"functional_group": "Shopping"
			// 		},
			// 		{
			// 			"category_code": 10364,
			// 			"category_name": "Catalogs",
			// 			"functional_group": "Shopping"
			// 		},
			// 		{
			// 			"category_code": 10365,
			// 			"category_name": "Contests \u0026 Surveys",
			// 			"functional_group": "Shopping"
			// 		},
			// 		{
			// 			"category_code": 10368,
			// 			"category_name": "Online Shopping",
			// 			"functional_group": "Shopping"
			// 		},
			// 		{
			// 			"category_code": 10367,
			// 			"category_name": "Engines",
			// 			"functional_group": "Shopping"
			// 		},
			// 		{
			// 			"category_code": 10369,
			// 			"category_name": "Product Reviews \u0026 Price Comparisons",
			// 			"functional_group": "Shopping"
			// 		},
			// 		{
			// 			"category_code": 10366,
			// 			"category_name": "Coupons",
			// 			"functional_group": "Shopping"
			// 		},
			// 		{
			// 			"category_code": 10370,
			// 			"category_name": "Shopping - Other",
			// 			"functional_group": "Shopping"
			// 		},
			// 		{
			// 			"category_code": 10371,
			// 			"category_name": "Auto Racing",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10372,
			// 			"category_name": "Baseball",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10373,
			// 			"category_name": "Bicycling",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10374,
			// 			"category_name": "Bodybuilding",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10375,
			// 			"category_name": "Boxing",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10376,
			// 			"category_name": "Canoeing/Kayaking",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10377,
			// 			"category_name": "Cheerleading",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10378,
			// 			"category_name": "Climbing",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10379,
			// 			"category_name": "Cricket",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10380,
			// 			"category_name": "Figure Skating",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10381,
			// 			"category_name": "Fly Fishing",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10382,
			// 			"category_name": "Football",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10383,
			// 			"category_name": "Freshwater Fishing",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10384,
			// 			"category_name": "Game \u0026 Fish",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10385,
			// 			"category_name": "Golf",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10386,
			// 			"category_name": "Horse Racing",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10387,
			// 			"category_name": "Horses",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10388,
			// 			"category_name": "Inline Skating",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10389,
			// 			"category_name": "Martial Arts",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10390,
			// 			"category_name": "Mountain Biking",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10391,
			// 			"category_name": "NASCAR Racing",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10392,
			// 			"category_name": "Olympics",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10415,
			// 			"category_name": "Sports - Other",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10393,
			// 			"category_name": "Paintball",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10394,
			// 			"category_name": "Power \u0026 Motorcycles",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10395,
			// 			"category_name": "Pro Basketball",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10396,
			// 			"category_name": "Pro Ice Hockey",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10397,
			// 			"category_name": "Rodeo",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10398,
			// 			"category_name": "Rugby",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10399,
			// 			"category_name": "Running/Jogging",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10400,
			// 			"category_name": "Sailing",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10401,
			// 			"category_name": "Saltwater Fishing",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10402,
			// 			"category_name": "Scuba Diving",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10403,
			// 			"category_name": "Skateboarding",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10404,
			// 			"category_name": "Skiing",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10405,
			// 			"category_name": "Snowboarding",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10406,
			// 			"category_name": "Sport Hunting",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10407,
			// 			"category_name": "Surfing/Bodyboarding",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10408,
			// 			"category_name": "Swimming",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10409,
			// 			"category_name": "Table Tennis/Ping-Pong",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10410,
			// 			"category_name": "Tennis",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10411,
			// 			"category_name": "Volleyball",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10412,
			// 			"category_name": "Walking",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10413,
			// 			"category_name": "Waterski/Wakeboard",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10414,
			// 			"category_name": "World Soccer",
			// 			"functional_group": "Sports"
			// 		},
			// 		{
			// 			"category_code": 10416,
			// 			"category_name": "3-D Graphics",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10417,
			// 			"category_name": "Animation",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10418,
			// 			"category_name": "Antivirus Software",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10419,
			// 			"category_name": "C/C++",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10420,
			// 			"category_name": "Cameras \u0026 Camcorders",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10421,
			// 			"category_name": "Computer Certification",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10422,
			// 			"category_name": "Computer Networking",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10423,
			// 			"category_name": "Computer Peripherals",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10424,
			// 			"category_name": "Computer Reviews",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10425,
			// 			"category_name": "Databases",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10426,
			// 			"category_name": "Desktop Publishing",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10427,
			// 			"category_name": "Desktop Video",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10429,
			// 			"category_name": "File Repositories",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10430,
			// 			"category_name": "Graphics Software",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10431,
			// 			"category_name": "Home Video/DVD",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10432,
			// 			"category_name": "Information Security",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10433,
			// 			"category_name": "Internet Phone \u0026 VOIP",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10434,
			// 			"category_name": "Internet Technology",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10435,
			// 			"category_name": "Java",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10436,
			// 			"category_name": "Javascript",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10437,
			// 			"category_name": "Linux",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10438,
			// 			"category_name": "Mac OS",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10460,
			// 			"category_name": "Technology - Other",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10439,
			// 			"category_name": "Mac Support",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10440,
			// 			"category_name": "Mobile Phones",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10441,
			// 			"category_name": "MP3/MIDI",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10442,
			// 			"category_name": "Net Conferencing",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10443,
			// 			"category_name": "Net for Beginners",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10444,
			// 			"category_name": "Network Security",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10446,
			// 			"category_name": "Palmtops/PDAs",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10447,
			// 			"category_name": "PC Support",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10448,
			// 			"category_name": "Peer-to-Peer",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10449,
			// 			"category_name": "Personal Storage",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10450,
			// 			"category_name": "Portable",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10428,
			// 			"category_name": "Entertainment",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10451,
			// 			"category_name": "Remote Access",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10453,
			// 			"category_name": "Unix",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10454,
			// 			"category_name": "Utilities",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10455,
			// 			"category_name": "Visual Basic",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10456,
			// 			"category_name": "Web Clip Art",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10457,
			// 			"category_name": "Web Design/HTML",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10458,
			// 			"category_name": "Web Hosting, ISP \u0026 Telco",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10459,
			// 			"category_name": "Windows",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10445,
			// 			"category_name": "Online Information Management",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10495,
			// 			"category_name": "APIs",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10496,
			// 			"category_name": "Internet of Things",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10493,
			// 			"category_name": "Blockchain",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10491,
			// 			"category_name": "Cryptocurrency",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10497,
			// 			"category_name": "A.I. \u0026 M.L.",
			// 			"functional_group": "Technology"
			// 		},
			// 		{
			// 			"category_code": 10461,
			// 			"category_name": "Adventure Travel",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10462,
			// 			"category_name": "Africa",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10463,
			// 			"category_name": "Air Travel",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10464,
			// 			"category_name": "Australia \u0026 New Zealand",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10465,
			// 			"category_name": "Bed \u0026 Breakfast",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10466,
			// 			"category_name": "Budget Travel",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10467,
			// 			"category_name": "Business Travel",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10468,
			// 			"category_name": "By US Locale",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10469,
			// 			"category_name": "Camping",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10470,
			// 			"category_name": "Canada",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10471,
			// 			"category_name": "Caribbean",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10472,
			// 			"category_name": "Cruises",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10473,
			// 			"category_name": "Eastern Europe",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10474,
			// 			"category_name": "Europe",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10489,
			// 			"category_name": "Travel - Other",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10475,
			// 			"category_name": "France",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10476,
			// 			"category_name": "Greece",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10477,
			// 			"category_name": "Honeymoons/Getaways",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10478,
			// 			"category_name": "Hotels",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10479,
			// 			"category_name": "Italy",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10480,
			// 			"category_name": "Japan",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10481,
			// 			"category_name": "Mexico \u0026 Central America",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10482,
			// 			"category_name": "National Parks",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10483,
			// 			"category_name": "Navigation",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10484,
			// 			"category_name": "South America",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10485,
			// 			"category_name": "Spas",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10486,
			// 			"category_name": "Theme Parks",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10487,
			// 			"category_name": "Traveling with Kids",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10488,
			// 			"category_name": "United Kingdom",
			// 			"functional_group": "Travel"
			// 		},
			// 		{
			// 			"category_code": 10700,
			// 			"category_name": "Uncategorized",
			// 			"functional_group": "Uncategorized"
			// 		},
			// 		{
			// 			"category_code": 10452,
			// 			"category_name": "Shareware \u0026 Freeware",
			// 			"functional_group": "Technology"
			// 		}
			// 	]
			// }

			// Threat Actors Metrics
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Threat Actor Metrics");
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Getting Threat Actor Information (This may take a moment)");
			if (isset($CubeJSResults['ThreatActors']['Body']->result)) {
				$ThreatActors = $CubeJSResults['ThreatActors']['Body']->result->data;
				// Workaround to EU / US Realm Alignment
				// if ($config['Realm'] == 'EU') {
				//   $ThreatActorInfo = $this->ThreatActors->GetB1ThreatActorsByIdEU($ThreatActors,$config['unnamed'],$config['substring'],$config['unknown']);
				// } elseif ($config['Realm'] == 'US') {
				  $ThreatActorInfo = $this->ThreatActors->GetB1ThreatActorsById($ThreatActors,$config['unnamed'],$config['substring'],$config['unknown']);
				// }
				if ($config['allTAInMetrics'] == 'true') {
					// switch($config['Realm']) {
					// 	case 'EU':
					// 		$ThreatActorsCountMetric = count(array_unique(array_column($ThreatActors, 'PortunusAggIPSummary.actor_id')));
					// 		break;
						// case 'US':
							$ThreatActorsCountMetric = count(array_unique(array_column($ThreatActors, 'ThreatActors.ikbactorid')));
							// break;
					// }
				} else {
					$ThreatActorsCountMetric = count($ThreatActorInfo);
				}
				$ThreatActorsCount = count($ThreatActorInfo);
				// End of workaround
	
				if (isset($ThreatActorInfo->error)) {
					$ThreatActorInfo = array();
					$ThreatActorSlideCount = 0;
				} else {
					$ThreatActorSlideCount = count($ThreatActorInfo);
				}
			} else {
				$ThreatActorsCount = 0;
			}
	
			// Threat Activity
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Threat Activity");
			$ThreatActivityEvents = $CubeJSResults['ThreatActivityEvents']['Body'];
			if (isset($ThreatActivityEvents->result->data[0])) {
				$ThreatActivityEventsCount = $ThreatActivityEvents->result->data[0]->{'PortunusAggInsight.threatCount'};
			} else {
				$ThreatActivityEventsCount = 0;
			}
	
			// DNS Firewall
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building DNS Firewall Activity");
			$DNSFirewallEvents = $CubeJSResults['DNSFirewallEvents']['Body'];
			if (isset($DNSFirewallEvents->result->data[0])) {
				$DNSFirewallEventsCount = $DNSFirewallEvents->result->data[0]->{'PortunusAggInsight.requests'};
			} else {
				$DNSFirewallEventsCount = 0;
			}
	
			// Device Count
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Device Count");
			$Devices = $CubeJSResults['Devices']['Body'];

			if (isset($Devices->result->data[0])) {
				$DeviceCount = $Devices->result->data[0]->{'PortunusAggInsight.deviceCount'};
			} else {
				$DeviceCount = 0;
			}
	
			// User Count
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building User Count");
			$Users = $CubeJSResults['Users']['Body'];
			if (isset($Users->result->data[0])) {
				$UserCount = $Users->result->data[0]->{'PortunusAggInsight.userCount'};
			} else {
				$UserCount = 0;
			}
	
			// Threat Insight Count
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Threat Insight Count");
			$ThreatInsight = $CubeJSResults['ThreatInsight']['Body'];
			if (isset($ThreatInsight->result->data)) {
				$ThreatInsightCount = count($ThreatInsight->result->data);
			} else {
				$ThreatInsightCount = 0;
			}
	
			// Threat View Count
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Threat View Count");
			$ThreatView = $CubeJSResults['ThreatView']['Body'];
			if (isset($ThreatView->result->data[0])) {
				$ThreatViewCount = $ThreatView->result->data[0]->{'PortunusAggInsight.tpropertyCount'};
			} else {
				$ThreatViewCount = 0;
			}
	
			// Source Count
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Sources Count");
			$Sources = $CubeJSResults['Sources']['Body'];
			if (isset($Sources->result->data[0])) {
				$SourcesCount = $Sources->result->data[0]->{'PortunusAggSecurity.networkCount'};
			} else {
				$SourcesCount = 0;
			}

			// Lookalike Threats
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Lookalike Threats");
			$LookalikeThreatCountUri = urlencode('/api/atclad/v1/lookalike_threat_counts?_filter=detected_at>="'.$StartDimension.'" and detected_at<="'.$EndDimension.'"');
			$LookalikeThreatCounts = $this->QueryCSP("get",$LookalikeThreatCountUri);

			// Zero Day DNS
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Zero Day DNS");
			// If Start and End date are < 7 days, use 7 days to ensure we get results. If more than 30 days, reduce to 30 days
			$ZDDStartDate = new DateTime($StartDimension);
			$ZDDEndDate = new DateTime($EndDimension);
			$DateDiff = $ZDDStartDate->diff($ZDDEndDate)->days;
			if ($DateDiff < 7) {
				$ZDDStartDimension = (new DateTime($EndDimension))->modify('-7 days')->format('Y-m-d\TH:i:s');
			}
			if ($DateDiff > 30) {
				$ZDDStartDimension = (new DateTime($EndDimension))->modify('-30 days')->format('Y-m-d\T').'00:00:00';
			}
			// Fix some weird behaviour with time
			$ZDDEndDimension = (new DateTime($EndDimension))->modify('+1 hour')->format('Y-m-d\TH:i:s');
			$ZeroDayDNSDetectionsUri = '/tide-rpz-stats/v1/zero_day_detections?from='.$ZDDStartDimension.'Z&to='.$ZDDEndDimension.'Z';
			$ZeroDayDNSDetections = $this->QueryCSP("get",$ZeroDayDNSDetectionsUri);

			// Work out number percentage of Suspicious and Malicious domains
			// If the property is not null and not Policy_ParkedDomain or Policy_NewlyObservedDomains then it is included in suspicious, unless it includes Phishing or Malware, then it's considered malicious 
			$ZeroDayDNSDetectionsTotalCount = 0;
			$ZeroDayDNSDetectionsSuspiciousCount = 0;
			$ZeroDayDNSDetectionsMaliciousCount = 0;
			if (isset($ZeroDayDNSDetections->detections) AND is_array($ZeroDayDNSDetections->detections)) {
				foreach ($ZeroDayDNSDetections->detections as $detection) {
					if (isset($detection->property) AND !in_array($detection->property, ['Policy_ParkedDomain','Policy_NewlyObservedDomains'])) {
						$ZeroDayDNSDetectionsSuspiciousCount++;
						if (stripos($detection->property,'Phishing') !== false OR stripos($detection->property,'Malware') !== false) {
							$ZeroDayDNSDetectionsMaliciousCount++;
						}
					}
					$ZeroDayDNSDetectionsTotalCount += $detection->count ?? 0; // Default to 0 if count is not set
				}
			}
			if ($ZeroDayDNSDetectionsTotalCount > 0) {
				$ZeroDayDNSDetectionsSuspiciousPercent = ($ZeroDayDNSDetectionsSuspiciousCount / $ZeroDayDNSDetectionsTotalCount) * 100;
				$ZeroDayDNSDetectionsMaliciousPercent = ($ZeroDayDNSDetectionsMaliciousCount / $ZeroDayDNSDetectionsTotalCount) * 100;
			} else {
				$ZeroDayDNSDetectionsSuspiciousPercent = 0;
				$ZeroDayDNSDetectionsMaliciousPercent = 0;
			}


			// New Loop for support of multiple selected templates
			// writeProgress result should be Selected Template 27 + 13 x N
			foreach ($SelectedTemplates as &$SelectedTemplate) {
				$embeddedDirectory = $SelectedTemplate['ExtractedDir'].'/ppt/embeddings/';
				$embeddedFiles = array_values(array_diff(scandir($embeddedDirectory), array('.', '..')));
				usort($embeddedFiles, 'strnatcmp');

				$this->logging->writeLog("Assessment","Embedded Files List","debug",['Template' => $SelectedTemplate, 'Embedded Files' => $embeddedFiles]);
	
				// Top detected properties
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Threat Properties");
				$TopDetectedProperties = $CubeJSResults['TopDetectedProperties']['Body'];
				if (isset($TopDetectedProperties->result->data)) {
					$EmbeddedTopDetectedProperties = getEmbeddedSheetFilePath('TopDetectedProperties', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets);
					$TopDetectedPropertiesSS = IOFactory::load($EmbeddedTopDetectedProperties);
					$RowNo = 2;
					foreach ($TopDetectedProperties->result->data as $TopDetectedProperty) {
						$TopDetectedPropertiesS = $TopDetectedPropertiesSS->getActiveSheet();
						$TopDetectedPropertiesS->setCellValue('A'.$RowNo, $TopDetectedProperty->{'PortunusDnsLogs.tproperty'});
						$TopDetectedPropertiesS->setCellValue('B'.$RowNo, $TopDetectedProperty->{'PortunusDnsLogs.tpropertyCount'});
						$RowNo++;
					}
					$TopDetectedPropertiesW = IOFactory::createWriter($TopDetectedPropertiesSS, 'Xlsx');
					$TopDetectedPropertiesW->save($EmbeddedTopDetectedProperties);
				}
		
				// Content filtration
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Content Filters");
				// $ContentFiltration = $CubeJSResults['ContentFiltration']['Body'];
				// Re-use High-Risk Websites data
				$ContentFiltration = $CubeJSResults['HighRiskWebsites']['Body'];
				if (isset($ContentFiltration->result->data)) {
					$EmbeddedContentFiltration = getEmbeddedSheetFilePath('ContentFiltration', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets);
					$ContentFiltrationSS = IOFactory::load($EmbeddedContentFiltration);
					$RowNo = 2;
					// Slice Array to limit size to 10
					$ContentFiltrationSliced = array_slice($ContentFiltration->result->data,0,10);
					foreach ($ContentFiltrationSliced as $ContentFilter) {
						$ContentFiltrationS = $ContentFiltrationSS->getActiveSheet();
						// $ContentFiltrationS->setCellValue('A'.$RowNo, $ContentFilter->{'PortunusAggWebcontent.category'});
						// $ContentFiltrationS->setCellValue('B'.$RowNo, $ContentFilter->{'PortunusAggWebcontent.categoryCount'});
						// Switch to using the Web Content Discovery APIs, rather than those called via the Dashboard.
						$ContentFiltrationS->setCellValue('A'.$RowNo, $ContentFilter->{'PortunusAggWebContentDiscovery.domain_category'});
						$ContentFiltrationS->setCellValue('B'.$RowNo, $ContentFilter->{'PortunusAggWebContentDiscovery.count'});
						$RowNo++;
					}
					$ContentFiltrationW = IOFactory::createWriter($ContentFiltrationSS, 'Xlsx');
					$ContentFiltrationW->save($EmbeddedContentFiltration);
				}

				// Traffic Analysis - DNS Activity
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building DNS Activity");
				$DNSActivityDaily = $CubeJSResults['DNSActivityDaily']['Body'];
				if (isset($DNSActivityDaily->result->data)) {
					$EmbeddedDNSActivityDaily = getEmbeddedSheetFilePath('DNSActivity', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets);
					$DNSActivityDailySS = IOFactory::load($EmbeddedDNSActivityDaily);
					$RowNo = 2;

					$DNSActivityDailyValues = array_map(function($item) {
						return $item->{'PortunusAggInsight.requests'};
					}, $DNSActivityDaily->result->data);
					// Calculate the average
					$DNSActivityDailySum = array_sum($DNSActivityDailyValues);
					$DNSActivityDailyCount = count($DNSActivityDailyValues);
					$DNSActivityDailyAverage = $DNSActivityDailyCount ? $DNSActivityDailySum / $DNSActivityDailyCount : 0;

					foreach ($DNSActivityDaily->result->data as $DNSActivityDay) {
						$DayTimestamp = new DateTime($DNSActivityDay->{'PortunusAggInsight.timestamp.day'});
						$DNSActivityDailyS = $DNSActivityDailySS->getActiveSheet();
						$DNSActivityDailyS->setCellValue('A'.$RowNo, $DayTimestamp->format('d/m/Y'));
						$DNSActivityDailyS->setCellValue('B'.$RowNo, $DNSActivityDay->{'PortunusAggInsight.requests'});
						$RowNo++;
					}
					$DNSActivityDailyW = IOFactory::createWriter($DNSActivityDailySS, 'Xlsx');
					$DNSActivityDailyW->save($EmbeddedDNSActivityDaily);
				}

				// Traffic Analysis - DNS Firewall Activity
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building DNS Firewall Activity");
				$DNSFirewallActivityDaily = $CubeJSResults['DNSFirewallActivityDaily']['Body'];
				if (isset($DNSFirewallActivityDaily->result->data)) {
					$EmbeddedDNSFirewallActivityDaily = getEmbeddedSheetFilePath('DNSFirewallActivity', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets);
					$DNSFirewallActivityDailySS = IOFactory::load($EmbeddedDNSFirewallActivityDaily);
					$RowNo = 2;

					$DNSFirewallActivityDailyValues = array_map(function($item) {
						return $item->{'PortunusAggSecurity.requests'};
					}, $DNSFirewallActivityDaily->result->data);
					// Calculate the average
					$DNSFirewallActivityDailySum = array_sum($DNSFirewallActivityDailyValues);
					$DNSFirewallActivityDailyCount = count($DNSFirewallActivityDailyValues);
					$DNSFirewallActivityDailyAverage = $DNSFirewallActivityDailyCount ? $DNSFirewallActivityDailySum / $DNSFirewallActivityDailyCount : 0;

					foreach ($DNSFirewallActivityDaily->result->data as $DNSFirewallActivityDay) {
						$DayTimestamp = new DateTime($DNSFirewallActivityDay->{'PortunusAggSecurity.timestamp.day'});
						$DNSFirewallActivityDailyS = $DNSFirewallActivityDailySS->getActiveSheet();
						$DNSFirewallActivityDailyS->setCellValue('A'.$RowNo, $DayTimestamp->format('d/m/Y'));
						$DNSFirewallActivityDailyS->setCellValue('B'.$RowNo, $DNSFirewallActivityDay->{'PortunusAggSecurity.requests'});
						$RowNo++;
					}
					$DNSFirewallActivityDailyW = IOFactory::createWriter($DNSFirewallActivityDailySS, 'Xlsx');
					$DNSFirewallActivityDailyW->save($EmbeddedDNSFirewallActivityDaily);
				}
		
				// Insight Distribution by Threat Type - Sheet 3
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building SOC Insight Threat Types");
				$InsightDistribution = $CubeJSResults['InsightDistribution']['Body'];
				if (isset($InsightDistribution->result->data)) {
					$EmbeddedInsightDistribution = getEmbeddedSheetFilePath('InsightDistribution', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets);
					$InsightDistributionSS = IOFactory::load($EmbeddedInsightDistribution);
					$RowNo = 2;
					foreach ($InsightDistribution->result->data as $InsightThreatType) {
						$InsightDistributionS = $InsightDistributionSS->getActiveSheet();
						$InsightDistributionS->setCellValue('A'.$RowNo, $InsightThreatType->{'InsightsAggregated.threatType'});
						$InsightDistributionS->setCellValue('B'.$RowNo, $InsightThreatType->{'InsightsAggregated.count'});
						$RowNo++;
					}
					$InsightDistributionW = IOFactory::createWriter($InsightDistributionSS, 'Xlsx');
					$InsightDistributionW->save($EmbeddedInsightDistribution);
				}
		
				// Threat Types (Lookalikes) - Sheet 4
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Lookalikes");
				if (isset($LookalikeThreatCounts->results)) {
					$EmbeddedLookalikes = getEmbeddedSheetFilePath('Lookalikes', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets);
					$LookalikeThreatCountsSS = IOFactory::load($EmbeddedLookalikes);
					$LookalikeThreatCountsS = $LookalikeThreatCountsSS->getActiveSheet();
					$RowNo = 2;
					if (isset($LookalikeThreatCounts->results->suspicious_count)) {
						$LookalikeThreatCountsS->setCellValue('A'.$RowNo, 'Suspicious');
						$LookalikeThreatCountsS->setCellValue('B'.$RowNo, $LookalikeThreatCounts->results->suspicious_count);
						$RowNo++;
					}
					if (isset($LookalikeThreatCounts->results->malware_count)) {
						$LookalikeThreatCountsS->setCellValue('A'.$RowNo, 'Malware');
						$LookalikeThreatCountsS->setCellValue('B'.$RowNo, $LookalikeThreatCounts->results->malware_count);
						$RowNo++;
					}
					if (isset($LookalikeThreatCounts->results->phishing_count)) {
						$LookalikeThreatCountsS->setCellValue('A'.$RowNo, 'Phishing');
						$LookalikeThreatCountsS->setCellValue('B'.$RowNo, $LookalikeThreatCounts->results->phishing_count);
						$RowNo++;
					}
					if (isset($LookalikeThreatCounts->results->others_count)) {
						$LookalikeThreatCountsS->setCellValue('A'.$RowNo, 'Others');
						$LookalikeThreatCountsS->setCellValue('B'.$RowNo, $LookalikeThreatCounts->results->others_count);
						$RowNo++;
					}
					$LookalikeThreatCountsW = IOFactory::createWriter($LookalikeThreatCountsSS, 'Xlsx');
					$LookalikeThreatCountsW->save($EmbeddedLookalikes);
				}

				// Application Detection - Slide 19
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Application Detection Page");
				$AppDiscoveryApplications->result->data = array_slice($AppDiscoveryApplications->result->data, 0, 10);
				$EmbeddedAppDiscovery = getEmbeddedSheetFilePath('AppDiscovery', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets);
				$AppDiscoverySS = IOFactory::load($EmbeddedAppDiscovery);
				$RowNo = 2;
				// Name, Category, Request Count, Status, Devices, Manufacturer
				if (isset($AppDiscoveryApplications->result->data)) {
					foreach ($AppDiscoveryApplications->result->data as $AppDiscovery) {
						$AppDiscoveryS = $AppDiscoverySS->getActiveSheet();
						$AppDiscoveryS->setCellValue('A'.$RowNo, $AppDiscovery->{'PortunusAggAppDiscovery.app_name'});
						$AppDiscoveryS->setCellValue('B'.$RowNo, $AppDiscovery->{'PortunusAggAppDiscovery.app_category'});
						$AppDiscoveryS->setCellValue('C'.$RowNo, $AppDiscovery->{'PortunusAggAppDiscovery.requests'});
						$AppDiscoveryS->setCellValue('D'.$RowNo, $AppDiscovery->{'PortunusAggAppDiscovery.app_approval'});
						$AppDiscoveryS->setCellValue('E'.$RowNo, $AppDiscovery->{'PortunusAggAppDiscovery.deviceCount'});
						$AppDiscoveryS->setCellValue('F'.$RowNo, $AppDiscovery->{'PortunusAggAppDiscovery.app_vendor'});
						$RowNo++;
					}
					$AppDiscoveryW = IOFactory::createWriter($AppDiscoverySS, 'Xlsx');
					$AppDiscoveryW->save($EmbeddedAppDiscovery);
				}

				// Web Content Discovery - Slide 20
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Web Content Discovery Page");
				$HighRiskWebsites->result->data = array_slice($HighRiskWebsites->result->data, 0, 10);
				$EmbeddedWebContentDiscovery = getEmbeddedSheetFilePath('WebContentDiscovery', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets);
				$WebContentDiscoverySS = IOFactory::load($EmbeddedWebContentDiscovery);
				$RowNo = 2;
				// Sub Category, Category, Request Count, Device Count
				if (isset($HighRiskWebsites->result->data)) {
					foreach ($HighRiskWebsites->result->data as $WebContentDiscovery) {

						// Determine Parent Category from Sub Category
						if (isset($WebContentCategories[$WebContentDiscovery->{'PortunusAggWebContentDiscovery.domain_category'}])) {
							$WebContentDiscoveryDomainCategory = $WebContentCategories[$WebContentDiscovery->{'PortunusAggWebContentDiscovery.domain_category'}]['Group'];
						} else {
							$WebContentDiscoveryDomainCategory = 'Uncategorized';
						}

						$WebContentDiscoveryS = $WebContentDiscoverySS->getActiveSheet();
						$WebContentDiscoveryS->setCellValue('A'.$RowNo, $WebContentDiscovery->{'PortunusAggWebContentDiscovery.domain_category'}); // Sub Category
						$WebContentDiscoveryS->setCellValue('B'.$RowNo, $WebContentDiscoveryDomainCategory); // Category
						$WebContentDiscoveryS->setCellValue('C'.$RowNo, $WebContentDiscovery->{'PortunusAggWebContentDiscovery.count'});
						$WebContentDiscoveryS->setCellValue('D'.$RowNo, $WebContentDiscovery->{'PortunusAggWebContentDiscovery.deviceCount'});
						$RowNo++;
					}
					$WebContentDiscoveryW = IOFactory::createWriter($WebContentDiscoverySS, 'Xlsx');
					$WebContentDiscoveryW->save($EmbeddedWebContentDiscovery);
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

				//
				// Do Threat Actor Stuff Here ....
				//
				// Skip Threat Actor Slides if Slide Number is set to 0
				if ($SelectedTemplate['ThreatActorSlide'] != 0) {
					$Progress = $this->writeProgress($config['UUID'],$Progress,"Generating Threat Actor Slides");
					// New slides to be appended after this slide number
					$ThreatActorSlideStart = $SelectedTemplate['ThreatActorSlide'];
					// Calculate the slide position based on above value
					$ThreatActorSlidePosition = $ThreatActorSlideStart-2;
		
					// Tag Numbers Start
					$TagStart = 100;
		
					// Create Document Fragments and set Starting Integers
					$xml_rels_f = $xml_rels->createDocumentFragment();
					$xml_rels_fstart = ($xml_rels->getElementsByTagName('Relationship')->length)+50;
					$xml_pres_f = $xml_pres->createDocumentFragment();
					$xml_pres_fstart = 14700;

					// Get Slide Count
					$SlidesCount = iterator_count(new FilesystemIterator($SelectedTemplate['ExtractedDir'].'/ppt/slides'));
					// Set first slide number
					$SlideNumber = $SlidesCount++;
					// Copy Blank Threat Actor Image
					copy($this->getDir()['PluginData'].'/images/logo-only.png',$SelectedTemplate['ExtractedDir'].'/ppt/media/logo-only.png');
					// Build new Threat Actor Slides & Update PPTX Resources
					$ThreatActorSlideCountIt = $ThreatActorSlideCount;
					if (isset($ThreatActorInfo)) {
						foreach  ($ThreatActorInfo as $TAI) {
							$KnownActor = $this->ThreatActors->getThreatActorConfigByName($TAI['actor_name']);
							if (($ThreatActorSlideCountIt - 1) > 0) {
								$xml_rels_f->appendXML('<Relationship Id="rId'.$xml_rels_fstart.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'.$SlideNumber.'.xml"/>');
								$xml_pres_f->appendXML('<p:sldId id="'.$xml_pres_fstart.'" r:id="rId'.$xml_rels_fstart.'"/>');
								$xml_rels_fstart++;
								$xml_pres_fstart++;
								copy($SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$ThreatActorSlideStart.'.xml',$SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$SlideNumber.'.xml');
								copy($SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$ThreatActorSlideStart.'.xml.rels',$SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$SlideNumber.'.xml.rels');
							} else {
								$SlideNumber = $ThreatActorSlideStart;
							}
							// Update Tag Numbers
							$TASFile = file_get_contents($SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$SlideNumber.'.xml');
							$TASFile = str_replace('#TATAG00', '#TATAG'.$TagStart, $TASFile);
							// Add Threat Actor Icon
							switch($SelectedTemplate['Orientation']) {
								case 'Portrait':
									$ThreatActorIconString = '<p:pic><p:nvPicPr><p:cNvPr id="36" name="Graphic 35"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{898E1A10-3ABF-AED0-2C71-1F26BBB6304B}"/></a:ext></a:extLst></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId115"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId115"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="5522998" y="2349624"/><a:ext cx="1246722" cy="1582377"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree>';
									$ThreatActorExternalLinkString = '<p:sp><p:nvSpPr><p:cNvPr id="7" name="Text Placeholder 20"><a:hlinkClick r:id="rId122"/><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{4A652F23-47D6-59A0-1D85-972482B29234}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr txBox="1"><a:spLocks/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="5574269" y="3869404"/><a:ext cx="1168604" cy="271567"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 20777"/></a:avLst></a:prstGeom><a:noFill/><a:ln w="19050"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln><a:effectLst><a:glow rad="63500"><a:srgbClr val="00B24C"><a:alpha val="40000"/></a:srgbClr></a:glow></a:effectLst></p:spPr><p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" anchorCtr="0"/><a:lstStyle><a:lvl1pPr marL="141755" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="620"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1736" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="425265" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1488" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="708774" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1240" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="992284" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1275794" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="1559303" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="1842813" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="2126323" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="2409833" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></a:lstStyle><a:p><a:pPr marL="0" indent="0" algn="ctr"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="300"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft><a:buClr><a:schemeClr val="tx1"/></a:buClr><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="600" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:latin typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="77"/><a:ea typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="0"/><a:cs typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="0"/></a:rPr><a:t>THREAT ACTOR REPORT</a:t></a:r><a:endParaRPr lang="en-US" sz="600" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:latin typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="77"/><a:ea typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="0"/><a:cs typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="0"/></a:endParaRPr></a:p></p:txBody></p:sp></p:spTree>';
									$VirusTotalString = '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="6" name="Straight Connector 5"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{3B07D3CE-83DF-306C-1740-B15E60D50B68}"/></a:ext></a:extLst></p:cNvPr><p:cNvCxnSpPr><a:cxnSpLocks/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="2663429" y="6816436"/><a:ext cx="0" cy="445863"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="9525"><a:solidFill><a:schemeClr val="accent3"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style></p:cxnSp><p:sp><p:nvSpPr><p:cNvPr id="11" name="Rectangle 10"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{5CF57A2B-9E16-9EF8-CC46-DEACEC1E9222}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="2390809" y="6646115"/><a:ext cx="546397" cy="151573"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" dirty="0" err="1"><a:solidFill><a:srgbClr val="101820"/></a:solidFill></a:endParaRPr></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="14" name="Graphic 13"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{6680C076-3929-2FD5-9B2D-C8EEC6FB5791}"/></a:ext></a:extLst></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId120"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId121"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="2407408" y="6670008"/><a:ext cx="499438" cy="100897"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic><p:sp><p:nvSpPr><p:cNvPr id="15" name="Oval 14"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{BF608D1F-2449-B2E7-8286-C23F058ABA75}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="2641147" y="7268668"/><a:ext cx="45719" cy="45719"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent3"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" dirty="0" err="1"><a:solidFill><a:srgbClr val="101820"/></a:solidFill></a:endParaRPr></a:p></p:txBody></p:sp></p:spTree>';
									break;
								case 'Landscape':
									$ThreatActorIconString = '<p:pic><p:nvPicPr><p:cNvPr id="36" name="Graphic 35"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{898E1A10-3ABF-AED0-2C71-1F26BBB6304B}"/></a:ext></a:extLst></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId115"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId115"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="4690800" y="2109600"/><a:ext cx="1246722" cy="1582377"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree>';
									$ThreatActorExternalLinkString = '<p:sp><p:nvSpPr><p:cNvPr id="7" name="Text Placeholder 20"><a:hlinkClick r:id="rId122"/><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{4A652F23-47D6-59A0-1D85-972482B29234}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr txBox="1"><a:spLocks/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="4755600" y="3618000"/><a:ext cx="1168604" cy="271567"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 20777"/></a:avLst></a:prstGeom><a:noFill/><a:ln w="19050"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln><a:effectLst><a:glow rad="63500"><a:srgbClr val="00B24C"><a:alpha val="40000"/></a:srgbClr></a:glow></a:effectLst></p:spPr><p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" anchorCtr="0"/><a:lstStyle><a:lvl1pPr marL="141755" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="620"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1736" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="425265" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1488" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="708774" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1240" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="992284" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1275794" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="1559303" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="1842813" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="2126323" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="2409833" indent="-141755" algn="l" defTabSz="567019" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="310"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="1116" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></a:lstStyle><a:p><a:pPr marL="0" indent="0" algn="ctr"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="300"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft><a:buClr><a:schemeClr val="tx1"/></a:buClr><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="600" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:latin typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="77"/><a:ea typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="0"/><a:cs typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="0"/></a:rPr><a:t>THREAT ACTOR REPORT</a:t></a:r><a:endParaRPr lang="en-US" sz="600" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:latin typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="77"/><a:ea typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="0"/><a:cs typeface="Lato" panose="020F0502020204030203" pitchFamily="34" charset="0"/></a:endParaRPr></a:p></p:txBody></p:sp></p:spTree>';
									$VirusTotalString = '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="6" name="Straight Connector 5"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{3B07D3CE-83DF-306C-1740-B15E60D50B68}"/></a:ext></a:extLst></p:cNvPr><p:cNvCxnSpPr><a:cxnSpLocks/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="8028000" y="2516400"/><a:ext cx="0" cy="445863"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="9525"><a:solidFill><a:schemeClr val="accent3"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style></p:cxnSp><p:sp><p:nvSpPr><p:cNvPr id="11" name="Rectangle 10"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{5CF57A2B-9E16-9EF8-CC46-DEACEC1E9222}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="7750800" y="2372400"/><a:ext cx="546397" cy="151573"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" dirty="0" err="1"><a:solidFill><a:srgbClr val="101820"/></a:solidFill></a:endParaRPr></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="14" name="Graphic 13"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{6680C076-3929-2FD5-9B2D-C8EEC6FB5791}"/></a:ext></a:extLst></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId120"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId121"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="7765200" y="2394000"/><a:ext cx="499438" cy="100897"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic><p:sp><p:nvSpPr><p:cNvPr id="15" name="Oval 14"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{BF608D1F-2449-B2E7-8286-C23F058ABA75}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="8006400" y="2966400"/><a:ext cx="45719" cy="45719"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent3"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" dirty="0" err="1"><a:solidFill><a:srgbClr val="101820"/></a:solidFill></a:endParaRPr></a:p></p:txBody></p:sp></p:spTree>';
									break;
							}
							$TASFile = str_replace('</p:spTree>',$ThreatActorIconString,$TASFile);
							// Append Virus Total Stuff if applicable to the slide
		
							// Workaround to EU / US Realm Alignment
							// if ($config['Realm'] == 'EU') {
							// 	if (isset($TAI['related_indicators_with_dates'])) {
							// 		foreach ($TAI['related_indicators_with_dates'] as $TAII) {
							// 			if (isset($TAII->vt_first_submission_date)) {
							// 				$TASFile = str_replace('</p:spTree>',$VirusTotalString,$TASFile);
							// 				$VTIndicatorFound = true;
							// 				break;
							// 			} else {
							// 				$VTIndicatorFound = false;
							// 			}
							// 		}
							// 	} else {
							// 		$VTIndicatorFound = false;
							// 	}
							// } elseif ($config['Realm'] == 'US') {
								if (isset($TAI['observed_iocs'])) {
									foreach ($TAI['observed_iocs'] as $TAII) {
										if (isset($TAII['ThreatActors.vtfirstdetectedts'])) {
											$TASFile = str_replace('</p:spTree>',$VirusTotalString,$TASFile);
											$VTIndicatorFound = true;
											break;
										} else {
											$VTIndicatorFound = false;
										}
									}
								} else {
									$VTIndicatorFound = false;
								}
							// }
		
							// Add Report Link
							// ** // Use the following code to link based on presence of 'infoblox_references' parameter
							// if (isset($TAI['infoblox_references'][0])) {
							// ** // Use the following code to link based on the Threat Actor config
							//$InfobloxReferenceFound = false;
							if ($KnownActor && $KnownActor['URLStub'] !== "") {
								$TASFile = str_replace('</p:spTree>',$ThreatActorExternalLinkString,$TASFile);
								//$InfobloxReferenceFound = true;
							}
							// Update Slide XML with changes
							file_put_contents($SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$SlideNumber.'.xml', $TASFile);
							$xml_tas = new DOMDocument('1.0', 'utf-8');
							$xml_tas->formatOutput = true;
							$xml_tas->preserveWhiteSpace = false;
							$xml_tas->load($SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$SlideNumber.'.xml.rels');
							foreach ($xml_tas->getElementsByTagName('Relationship') as $element) {
								// Remove notes references to avoid having to create unneccessary notes resources
								if ($element->getAttribute('Type') == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide") {
									$element->remove();
								}
							}
							$xml_tas_f = $xml_tas->createDocumentFragment();
							if ($KnownActor && $KnownActor['PNG'] !== "" && $KnownActor['SVG'] !== "") {
								$SVG = $KnownActor['SVG'];
								$PNG = $KnownActor['PNG'];
								// Threat Actor PNG
								copy($this->getDir()['Assets'].'/images/Threat Actors/Uploads/'.$PNG,$SelectedTemplate['ExtractedDir'].'/ppt/media/'.$PNG);
								copy($this->getDir()['Assets'].'/images/Threat Actors/Uploads/'.$SVG,$SelectedTemplate['ExtractedDir'].'/ppt/media/'.$SVG);
								$xml_tas_f->appendXML('<Relationship Id="rId115" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/'.$PNG.'"/>');
								$xml_tas_f->appendXML('<Relationship Id="rId116" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/'.$SVG.'"/>');
							} else {
								$xml_tas_f->appendXML('<Relationship Id="rId115" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo-only.png"/>');
								$xml_tas_f->appendXML('<Relationship Id="rId116" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo-only.svg"/>');
							}
		
							// Virus Total PNG / SVG
							if ($VTIndicatorFound) {
								copy($this->getDir()['PluginData'].'/images/virustotal.png',$SelectedTemplate['ExtractedDir'].'/ppt/media/virustotal.png');
								copy($this->getDir()['PluginData'].'/images/virustotal.svg',$SelectedTemplate['ExtractedDir'].'/ppt/media/virustotal.svg');
								$xml_tas_f->appendXML('<Relationship Id="rId120" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/virustotal.png"/>');
								$xml_tas_f->appendXML('<Relationship Id="rId121" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/virustotal.svg"/>');
							}
		
							// Infoblox Blog URL
							// ** // Use the following code to link based on presence of 'infoblox_references' parameter
							// if ($InfobloxReferenceFound) {
							//     $xml_tas_f2 = $xml_tas->createDocumentFragment();
							//     $xml_tas_f2->appendXML('<Relationship Id="rId122" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="'.$TAI['infoblox_references'][0].'" TargetMode="External"/>');
							//     $xml_tas->getElementsByTagName('Relationships')->item(0)->appendChild($xml_tas_f2);
							// }
							// ** // Use the following code to link based on the Threat Actor config
							if ($KnownActor && $KnownActor['URLStub'] !== "") {
								$URL = $KnownActor['URLStub'];
								$xml_tas_f->appendXML('<Relationship Id="rId122" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="'.$URL.'" TargetMode="External"/>');
							}
		
							$xml_tas->getElementsByTagName('Relationships')->item(0)->appendChild($xml_tas_f);
							$xml_tas->save($SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$SlideNumber.'.xml.rels');
							$TagStart += 10;
							// Iterate slide number
							$SlideNumber++;
							$ThreatActorSlideCountIt--;
						}
		
						// Append Elements to Core XML Files
						$xml_rels->getElementsByTagName('Relationships')->item(0)->appendChild($xml_rels_f);
						// Append new slides to specific position
						$xml_pres->getElementsByTagName('sldId')->item($ThreatActorSlidePosition)->after($xml_pres_f);

						//
						// End of Threat Actors
						//
					}
				} else {
					$Progress = $this->writeProgress($config['UUID'],$Progress,"Skipping Threat Actor Slides");
				}

				// Save Core XML Files
				$xml_rels->save($SelectedTemplate['ExtractedDir'].'/ppt/_rels/presentation.xml.rels');
				$xml_pres->save($SelectedTemplate['ExtractedDir'].'/ppt/presentation.xml');

				// Rebuild Powerpoint Template Zip(s)
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Stitching Powerpoint Template(s)");
				compressZip($this->getDir()['Files'].'/reports/report'.'-'.$config['UUID'].'-'.$SelectedTemplate['FileName'],$SelectedTemplate['ExtractedDir']);

				// Cleanup Extracted Zip(s)
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Cleaning up");
				// rmdirRecursive($SelectedTemplate['ExtractedDir']);

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
				$mapping = replaceTag($mapping,'#CUSTOMER',$CurrentAccount->name);
				$mapping = replaceTag($mapping,'#DATE',date("jS F Y"));
				$StartDate = new DateTime($StartDimension);
				$EndDate = new DateTime($EndDimension);
				$mapping = replaceTag($mapping,'#DATESOFCOLLECTION',$StartDate->format("jS F Y").' - '.$EndDate->format("jS F Y"));
				$mapping = replaceTag($mapping,'#NAME',$UserInfo->result->name);
				$mapping = replaceTag($mapping,'#EMAIL',$UserInfo->result->email);
		
				$PLACEHOLDER = 0; // Placeholder for future metrics

				##// Slide 5 - Executive Summary
				$mapping = replaceTag($mapping,'#TAG01',number_abbr($HighEventsCount)); // High-Risk Events
				$mapping = replaceTag($mapping,'#TAG02',number_abbr($HighRiskWebsiteCount)); // High-Risk Websites
				$mapping = replaceTag($mapping,'#TAG03',number_abbr($DataExfilEventsCount)); // Data Exfil / Tunneling
				$mapping = replaceTag($mapping,'#TAG04',number_abbr($LookalikeThreatCount)); // Lookalike Domains
				$mapping = replaceTag($mapping,'#TAG05',number_abbr($ZeroDayDNSEventsCount)); // Zero Day DNS

				// ** TODO ** //
				$mapping = replaceTag($mapping,'#TAG06',number_abbr($SuspiciousEventsCount)); // Suspicious Domains ???? HIGH RISK ??
				$mapping = replaceTag($mapping,'#TAG07',number_abbr($FirstToDetectTotalDomains)); // First to Detect (Domain Count)
				$mapping = replaceTag($mapping,'#TAG08',$TotalBandwidthSaved); // Bandwidth Savings
		
				##// Slide 6 - Security Indicator Summary
				$mapping = replaceTag($mapping,'#TAG09',number_abbr($DNSActivityCount)); // DNS Requests
				$mapping = replaceTag($mapping,'#TAG10',number_abbr($HighEventsCount)); // High-Risk Events
				$mapping = replaceTag($mapping,'#TAG11',number_abbr($MediumEventsCount)); // Medium-Risk Events
				$mapping = replaceTag($mapping,'#TAG12',number_abbr($TotalInsights)); // Insights
				$mapping = replaceTag($mapping,'#TAG13',number_abbr($LookalikeThreatCount)); // Custom Lookalike Domains
				$mapping = replaceTag($mapping,'#TAG14',number_abbr($DOHEventsCount)); // DoH
				$mapping = replaceTag($mapping,'#TAG15',number_abbr($ZeroDayDNSEventsCount)); // Zero Day DNS

				// ** TODO ** //
				$mapping = replaceTag($mapping,'#TAG16',number_abbr($SuspiciousEventsCount)); // Suspicious Domains ???? HIGH RISK ??
		
				$mapping = replaceTag($mapping,'#TAG17',number_abbr($NODEventsCount)); // Newly Observed Domains
				$mapping = replaceTag($mapping,'#TAG18',number_abbr($DGAEventsCount)); // Domain Generated Algorithms
				$mapping = replaceTag($mapping,'#TAG19',number_abbr($DataExfilEventsCount)); // DNS Tunnelling
				$mapping = replaceTag($mapping,'#TAG20',number_abbr($AppDiscoveryApplicationsCount)); // Unique Applications
				$mapping = replaceTag($mapping,'#TAG21',number_abbr($HighRiskWebCategoryCount)); // High-Risk Web Categories
				$mapping = replaceTag($mapping,'#TAG22',number_abbr($ThreatActorsCountMetric)); // Threat Actors
				$mapping = replaceTag($mapping,'#TAG23',number_abbr($FirstToDetectTotalDomains)); // First to Detect (Domain Count)
				$mapping = replaceTag($mapping,'#TAG24',number_abbr($MaliciousTDSCounts['Domains'])); // Malicious TDS Events

				// Only present on landscape template
				$mapping = replaceTag($mapping,'#TAG25',number_abbr($DNSFirewallActivityDailyAverage)); // Avg Events Per Day

				##// Slide 7 - Bandwidth Savings
				$mapping = replaceTag($mapping,'#TAG26',$TotalBandwidthSaved); // Total Savings (MB/GB/TB)
				$mapping = replaceTag($mapping,'#TAG27',$BandwidthSavedPercentage); // Percentage Overall %
	
				##// Slide 8 - Additional Threat Intel Insights
				// ** CURRENTLY DONE MANUALLY BY ITI ** //
		
				##// Slide 11 - Traffic Usage Analysis
				// Total DNS Activity
				$mapping = replaceTag($mapping,'#TAG28',number_abbr($DNSActivityCount));
				// DNS Firewall Activity
				$mapping = replaceTag($mapping,'#TAG29',number_abbr($HML)); // Total
				$mapping = replaceTag($mapping,'#TAG30',number_abbr($HighEventsCount)); // High Int
				$mapping = replaceTag($mapping,'#TAG31',number_format($HighPerc,2).'%'); // High Percent
				$mapping = replaceTag($mapping,'#TAG32',number_abbr($MediumEventsCount)); // Medium Int
				$mapping = replaceTag($mapping,'#TAG33',number_format($MediumPerc,2).'%'); // Medium Percent
				$mapping = replaceTag($mapping,'#TAG34',number_abbr($LowEventsCount)); // Low Int
				$mapping = replaceTag($mapping,'#TAG35',number_format($LowPerc,2).'%'); // Low Percent
				// Threat Activity
				$mapping = replaceTag($mapping,'#TAG36',number_abbr($ThreatActivityEventsCount));
				// Data Exfiltration Incidents
				$mapping = replaceTag($mapping,'#TAG37',number_abbr($DataExfilEventsCount));
		
				##// Slide 12 - Traffic Analysis - DNS Activity
				$mapping = replaceTag($mapping,'#TAG38',number_abbr($DNSActivityDailyAverage));
				##// Slide 13 - Traffic Analysis - DNS Firewall Activity
				$mapping = replaceTag($mapping,'#TAG39',number_abbr($DNSFirewallActivityDailyAverage));
	
				##// Slide 15 - Key Insights
				// Insight Severity
				$mapping = replaceTag($mapping,'#TAG40',number_abbr($TotalInsights)); // Total Open Insights
				$mapping = replaceTag($mapping,'#TAG41',number_abbr($CriticalInsights)); // Critical Priority Insights
				$mapping = replaceTag($mapping,'#TAG42',number_abbr($HighInsights)); // High Priority Insights
				$mapping = replaceTag($mapping,'#TAG43',number_abbr($MediumInsights)); // Medium Priority Insights
				$mapping = replaceTag($mapping,'#TAG44',number_abbr($LowInsights)); // Low Priority Insights
				$mapping = replaceTag($mapping,'#TAG45',number_abbr($InfoInsights)); // Info Priority Insights
				// Event To Insight Aggregation
				$mapping = replaceTag($mapping,'#TAG46',number_abbr($SecurityEventsCount)); // Events
				$mapping = replaceTag($mapping,'#TAG47',number_abbr($TotalInsights)); // Key Insights
		
				##// Slide 19 - Application Detection
				$mapping = replaceTag($mapping,'#TAG48',number_abbr($AppDiscoveryTotalRequestsCount)); // Total Requests
				$mapping = replaceTag($mapping,'#TAG49',number_abbr($AppDiscoveryTotalDevicesCount)); // Total Devices

				##// Slide 19 - Web Content Discovery
				$mapping = replaceTag($mapping,'#TAG50',number_abbr($WebContentTotalRequestsCount)); // Total Requests
				$mapping = replaceTag($mapping,'#TAG51',number_abbr($WebContentTotalDevicesCount)); // Total Devices

				##// Slide 24 - Lookalike Domains
				$mapping = replaceTag($mapping,'#TAG52',number_abbr($LookalikeTotalCount)); // Total Lookalikes
				// if ($LookalikeTotalPercentage >= 0){$arrow='↑';} else {$arrow='↓';}
				// $mapping = replaceTag($mapping,'#TAG39',$arrow); // Arrow Up/Down
				// $mapping = replaceTag($mapping,'#TAG40',number_abbr($LookalikeTotalPercentage)); // Total Percentage Increase
				$mapping = replaceTag($mapping,'#TAG53',number_abbr($LookalikeCustomCount)); // Total Lookalikes from Custom Watched Domains
				// if ($LookalikeCustomPercentage >= 0){$arrow='↑';} else {$arrow='↓';}
				// $mapping = replaceTag($mapping,'#TAG42',$arrow); // Arrow Up/Down
				// $mapping = replaceTag($mapping,'#TAG43',number_abbr($LookalikeCustomPercentage)); // Custom Percentage Increase
				$mapping = replaceTag($mapping,'#TAG54',number_abbr($LookalikeThreatCount)); // Threats from Custom Watched Domains
				// if ($LookalikeThreatPercentage >= 0){$arrow='↑';} else {$arrow='↓';}
				// $mapping = replaceTag($mapping,'#TAG45',$arrow); // Arrow Up/Down
				// $mapping = replaceTag($mapping,'#TAG46',number_abbr($LookalikeThreatPercentage)); // Threats Percentage Increase
		
				##// Slide 29/31 - Security Activities
				$mapping = replaceTag($mapping,'#TAG55',number_abbr($SecurityEventsCount)); // Security Events
				$mapping = replaceTag($mapping,'#TAG56',number_abbr($DNSFirewallEventsCount)); // DNS Firewall
				$mapping = replaceTag($mapping,'#TAG57',number_abbr($WebContentSecurityEventsCount)); // Web Content
				$mapping = replaceTag($mapping,'#TAG58',number_abbr($DeviceCount)); // Devices
				$mapping = replaceTag($mapping,'#TAG59',number_abbr($UserCount)); // Users
				$mapping = replaceTag($mapping,'#TAG60',number_abbr($TotalInsights)); // Insights
				$mapping = replaceTag($mapping,'#TAG61',number_abbr($ThreatInsightCount)); // Threat Insight
				$mapping = replaceTag($mapping,'#TAG62',number_abbr($ThreatViewCount)); // Threat View
				$mapping = replaceTag($mapping,'#TAG63',number_abbr($SourcesCount)); // Sources

				##// Slide 35/38 - Zero Day DNS
				$mapping = replaceTag($mapping,'#TAG64',number_abbr($ZeroDayDNSDetectionsTotalCount)); // Zero Day DNS Events // ZeroDayDNSEventsCount
				$mapping = replaceTag($mapping,'#TAG65',number_abbr($ZeroDayDNSDetectionsSuspiciousPercent).'%'); // Suspicious Events // ZeroDayDNSSuspiciousEventsCount
				$mapping = replaceTag($mapping,'#TAG66',number_abbr($ZeroDayDNSDetectionsMaliciousPercent).'%'); // Malicious Events // ZeroDayDNSMaliciousEventsCount

				##// Slide 32/34 - Threat Actors
				// This is where the Threat Actor Tag replacement occurs
				// Set Tag Start Number
				$TagStart = 100;
				if (isset($ThreatActorInfo)) {
					foreach ($ThreatActorInfo as $TAI) {
						// Workaround for EU / US Realm Alignment
						// if ($config['Realm'] == 'EU') {
						// 	// Get sorted list of observed IOCs not found in Virus Total
						// 	if (isset($TAI['related_indicators_with_dates'])) {
						// 		$ObservedIndicators = $TAI['related_indicators_with_dates'];
						// 		$IndicatorCount = count($TAI['related_indicators_with_dates']);
						// 		$IndicatorsInVT = [];
						// 		if ($IndicatorCount > 0) {
						// 			foreach ($ObservedIndicators as $OI) {
						// 				if (array_key_exists('vt_first_submission_date', json_decode(json_encode($OI), true))) {
						// 					$IndicatorsInVT[] = $OI;
						// 				}
						// 			}
						// 		}
						// 		if (count($IndicatorsInVT) > 0) {
						// 			// Sort the array based on the time difference
						// 			usort($IndicatorsInVT, function($a, $b) {
						// 				return $this->calculateVirusTotalDifference($b) <=> $this->calculateVirusTotalDifference($a);
						// 			});
						// 			$IndicatorsNotObserved = $TAI['related_count'] - count($ObservedIndicators);
						// 			$ExampleDomain = $IndicatorsInVT[0]->indicator;
						// 			$FirstSeen = new DateTime($IndicatorsInVT[0]->te_ik_submitted);
						// 			$LastSeen = new DateTime($IndicatorsInVT[0]->te_customer_last_dns_query);
						// 			$VTDate = new DateTime($IndicatorsInVT[0]->vt_first_submission_date);
						// 			$ProtectedFor = $FirstSeen->diff($LastSeen)->days;
						// 			$DaysAhead = 'Discovered '.($ProtectedFor - $LastSeen->diff($VTDate)->days).' days ahead';
						// 		} else {
						// 			$IndicatorsNotObserved = $TAI['related_count'] - count($ObservedIndicators);
						// 			$ExampleDomain = $ObservedIndicators[0]->indicator;
						// 			$FirstSeen = new DateTime($ObservedIndicators[0]->te_ik_submitted);
						// 			$LastSeen = new DateTime($ObservedIndicators[0]->te_customer_last_dns_query);
						// 			$DaysAhead = 'Discovered';
						// 			$ProtectedFor = $FirstSeen->diff($LastSeen)->days;
						// 		}
						// 	} else {
						// 		$IndicatorsNotObserved = 'N/A';
						// 		$ExampleDomain = 'N/A';
						// 		$FirstSeen = new DateTime('1901-01-01 00:00');
						// 		$LastSeen = new DateTime('1901-01-01 00:00');
						// 		$DaysAhead = 'Discovered';
						// 		$ProtectedFor = 'N/A';
						// 		$IndicatorCount = 'N/A';
						// 	}
						// } elseif ($config['Realm'] == 'US') {
							if (isset($TAI['observed_iocs'])) {
								$ObservedIndicators = $TAI['observed_iocs'];
								$IndicatorCount = $TAI['observed_count'];
								$IndicatorsInVT = [];
								if ($IndicatorCount > 0) {
									foreach ($ObservedIndicators as $OI) {
										if (array_key_exists('ThreatActors.vtfirstdetectedts', json_decode(json_encode($OI), true))) {
											$IndicatorsInVT[] = $OI;
										}
									}
								}
								if (count($IndicatorsInVT) > 0) {
									// Sort the array based on the time difference
									usort($IndicatorsInVT, function($a, $b) {
										return $this->calculateVirusTotalDifference($b) <=> $this->calculateVirusTotalDifference($a);
									});
									$IndicatorsNotObserved = $TAI['related_count'] - $IndicatorCount;
									$ExampleDomain = $IndicatorsInVT[0]['ThreatActors.domain'];
									$FirstSeen = new DateTime($IndicatorsInVT[0]['ThreatActors.ikbfirstsubmittedts']);
									$LastSeen = new DateTime($IndicatorsInVT[0]['ThreatActors.lastdetectedts']);
									$VTDate = new DateTime($IndicatorsInVT[0]['ThreatActors.vtfirstdetectedts']);
									$ProtectedFor = $FirstSeen->diff($LastSeen)->days;
									$DaysAhead = 'Discovered '.($ProtectedFor - $LastSeen->diff($VTDate)->days).' days ahead';
								} else {
									$IndicatorsNotObserved = $TAI['related_count'];
									$ExampleDomain = $ObservedIndicators[0]['ThreatActors.domain'];
									$FirstSeen = new DateTime($ObservedIndicators[0]['ThreatActors.ikbfirstsubmittedts']);
									$LastSeen = new DateTime($ObservedIndicators[0]['ThreatActors.lastdetectedts']);
									$DaysAhead = 'Discovered';
									$ProtectedFor = $FirstSeen->diff($LastSeen)->days;
								}
							} else {
								$IndicatorsNotObserved = 'N/A';
								$ExampleDomain = 'N/A';
								$FirstSeen = new DateTime('1901-01-01 00:00');
								$LastSeen = new DateTime('1901-01-01 00:00');
								$DaysAhead = 'Discovered';
								$ProtectedFor = 'N/A';
								$IndicatorCount = 'N/A';
							}
						// }
						// Workaround End
		
						$mapping = replaceTag($mapping,'#TATAG'.$TagStart.'01',ucwords($TAI['actor_name'])); // Threat Actor Name
						$mapping = replaceTag($mapping,'#TATAG'.$TagStart.'02',$TAI['actor_description']); // Threat Actor Description
						$mapping = replaceTag($mapping,'#TATAG'.$TagStart.'03',$IndicatorCount); // Number of Observed IOCs
						$mapping = replaceTag($mapping,'#TATAG'.$TagStart.'04',$IndicatorsNotObserved); // Number of Observed IOCs not observed
						$mapping = replaceTag($mapping,'#TATAG'.$TagStart.'05',$TAI['related_count']); // Number of Related IOCs
						$mapping = replaceTag($mapping,'#TATAG'.$TagStart.'06',$DaysAhead); // Discovered X Days Ahead
						$mapping = replaceTag($mapping,'#TATAG'.$TagStart.'07',$FirstSeen->format('d/m/y')); // First Detection Date
						$mapping = replaceTag($mapping,'#TATAG'.$TagStart.'08',$LastSeen->format('d/m/y')); // Last Detection Date
						$mapping = replaceTag($mapping,'#TATAG'.$TagStart.'09',$ProtectedFor); // Protected X Days
						$mapping = replaceTag($mapping,'#TATAG'.$TagStart.'10',$ExampleDomain); // Example Domain
						$TagStart += 10;
		
					}
				}

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
	
			// Write to Anonymised Metrics
			// dns_requests INTEGER,
			// security_events_high_risk INTEGER,
			// security_events_medium_risk INTEGER,
			// security_events_low_risk INTEGER,
			// security_events_doh INTEGER,
			// security_events_zero_day INTEGER,
			// security_events_suspicious INTEGER,
			// security_events_newly_observed_domains INTEGER,
			// security_events_dga INTEGER,
			// security_events_tunnelling INTEGER,
			// security_insights INTEGER,
			// security_threat_actors INTEGER,
			// web_unique_applications INTEGER,
			// web_high_risk_categories INTEGER,
			// lookalikes_custom_domains INTEGER
			$this->AssessmentReporting->newSecurityMetricsEntry([
				'date_start' => $StartDate,
				'date_end' => $EndDate,
				'dns_requests' => $DNSActivityCount,
				'security_events_high_risk' => $HighEventsCount,
				'security_events_medium_risk' => $MediumEventsCount,
				'security_events_low_risk' => $LowEventsCount,
				'security_events_doh' => $DOHEventsCount,
				'security_events_zero_day' => $ZeroDayDNSEventsCount,
				'security_events_suspicious' => $SuspiciousEventsCount,
				'security_events_newly_observed_domains' => $NODEventsCount,
				'security_events_dga' => $DGAEventsCount,
				'security_events_tunnelling' => $DataExfilEventsCount,
				'security_insights' => $TotalInsights,
				'security_threat_actors' => $ThreatActorsCountMetric,
				'web_unique_applications' => $AppDiscoveryApplicationsCount,
				'web_high_risk_categories' => $HighRiskWebCategoryCount,
				'lookalikes_custom_domains' => $LookalikeThreatCount
			]);

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
			'Total' => ($Total * 15) + 32,
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
	
	private function calculateProtectedDifference($te_ik_submitted,$te_customer_last_dns_query) {
		$submitted = new DateTime($te_ik_submitted);
		$lastQuery = new DateTime($te_customer_last_dns_query);
		return $lastQuery->getTimestamp() - $submitted->getTimestamp();
	}
	
	private function calculateVirusTotalDifference($obj) {
		// Workaround for EU / US Realm Alignment
		// if ($this->Realm == 'EU') {
		// 	$submitted = new DateTime($obj->te_ik_submitted);
		// 	$vtsubmitted = new DateTime($obj->vt_first_submission_date);
		// } elseif ($this->Realm == 'US') {
			$submitted = new DateTime($obj['ThreatActors.ikbfirstsubmittedts']);
			$vtsubmitted = new DateTime($obj['ThreatActors.vtfirstdetectedts']);
		// }
		return $vtsubmitted->getTimestamp() - $submitted->getTimestamp();
	}
}