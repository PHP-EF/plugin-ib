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
				'ZeroDayDNSEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2","3"]},{"member":"PortunusAggInsight.tclass","operator":"equals","values":["Zero Day DNS"]}],"ungrouped":false}',
				'SuspiciousEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2"]},{"member":"PortunusAggInsight.tclass","operator":"equals","values":["Suspicious"]}],"ungrouped":false}',
				'HighRiskWebsites' => '{"timeDimensions":[{"dimension":"PortunusAggWebContentDiscovery.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"measures":["PortunusAggWebContentDiscovery.count","PortunusAggWebContentDiscovery.deviceCount"],"dimensions":["PortunusAggWebContentDiscovery.domain_category"],"order":{"PortunusAggWebContentDiscovery.count":"desc"},"filters":[{"member":"PortunusAggWebContentDiscovery.domain_category","operator":"equals","values":["'.$HighRiskCategoryList.'"]}]}',
				'DOHEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2"]},{"member":"PortunusAggInsight.tproperty","operator":"equals","values":["DoHService"]}],"ungrouped":false}',
				'NODEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2"]},{"member":"PortunusAggInsight.tproperty","operator":"equals","values":["NewlyObservedDomains"]}],"ungrouped":false}',
				'DGAEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"or":[{"member":"PortunusAggInsight.tproperty","operator":"equals","values":["suspicious_rdga","suspicious_dga","DGA"]},{"member":"PortunusAggInsight.tclass","operator":"equals","values":["DGA","MalwareC2DGA"]}]},{"member":"PortunusAggInsight.type","operator":"equals","values":["2","3"]}],"ungrouped":false}',
				'UniqueApplications' => '{"measures":["PortunusAggAppDiscovery.requests"],"dimensions":["PortunusAggAppDiscovery.app_name","PortunusAggAppDiscovery.app_approval"],"timeDimensions":[{"dimension":"PortunusAggAppDiscovery.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggAppDiscovery.app_name","operator":"set"},{"member":"PortunusAggAppDiscovery.app_name","operator":"notEquals","values":[""]}],"order":{}}',
				'ThreatActivityEvents' => '{"measures":["PortunusAggInsight.threatCount"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2"]},{"member":"PortunusAggInsight.severity","operator":"equals","values":["High","Medium","Low"]},{"member":"PortunusAggInsight.threat_indicator","operator":"notEquals","values":[""]}],"limit":"1","ungrouped":false}',
				'DNSFirewallEvents' => '{"measures":["PortunusAggInsight.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggInsight.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"and":[{"member":"PortunusAggInsight.type","operator":"equals","values":["2"]},{"or":[{"member":"PortunusAggInsight.severity","operator":"equals","values":["High","Medium","Low"]},{"and":[{"member":"PortunusAggInsight.severity","operator":"equals","values":["Info"]},{"member":"PortunusAggInsight.policy_action","operator":"equals","values":["Block","Log"]}]}]},{"member":"PortunusAggInsight.confidence","operator":"equals","values":["High","Medium","Low"]}]}],"limit":"1","ungrouped":false}',
				'WebContentEvents' => '{"measures":["PortunusAggWebcontent.requests"],"dimensions":[],"timeDimensions":[{"dimension":"PortunusAggWebcontent.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"filters":[{"member":"PortunusAggWebcontent.type","operator":"equals","values":["3"]},{"member":"PortunusAggWebcontent.category","operator":"notEquals","values":[null]}],"limit":"1","ungrouped":false}',
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
				'TopDetectedProperties' => 4,
				'ContentFiltration' => 5,
				'DNSActivity' => 6,
				'DNSFirewallActivity' => 7,
				'InsightDistribution' => 8,
				'Lookalikes' => 9
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
	
			// Unique Applications
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building list of Unique Applications");
			$UniqueApplications = $CubeJSResults['UniqueApplications']['Body'];
			if (isset($UniqueApplications->result->data)) {
				$UniqueApplicationsCount = count($UniqueApplications->result->data);
			} else {
				$UniqueApplicationsCount = 0;
			}
	
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
	
			// Web Content
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Web Content Events");
			$WebContentEvents = $CubeJSResults['WebContentEvents']['Body'];
			if (isset($WebContentEvents->result->data[0])) {
				$WebContentEventsCount = $WebContentEvents->result->data[0]->{'PortunusAggWebcontent.requests'};
			} else {
				$WebContentEventsCount = 0;
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

			// New Loop for support of multiple selected templates
			// writeProgress result should be Selected Template 27 + 13 x N
			foreach ($SelectedTemplates as &$SelectedTemplate) {
				$embeddedDirectory = $SelectedTemplate['ExtractedDir'].'/ppt/embeddings/';
				$embeddedFiles = array_values(array_diff(scandir($embeddedDirectory), array('.', '..')));
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
				rmdirRecursive($SelectedTemplate['ExtractedDir']);

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
				$mapping = replaceTag($mapping,'#TAG502',number_abbr($HighEventsCount)); // High-Risk Events
				$mapping = replaceTag($mapping,'#TAG503',number_abbr($HighRiskWebsiteCount)); // High-Risk Websites
				$mapping = replaceTag($mapping,'#TAG504',number_abbr($DataExfilEventsCount)); // Data Exfil / Tunneling
				$mapping = replaceTag($mapping,'#TAG505',number_abbr($LookalikeThreatCount)); // Lookalike Domains
				$mapping = replaceTag($mapping,'#TAG506',number_abbr($ZeroDayDNSEventsCount)); // Zero Day DNS

				// ** TODO ** //
				$mapping = replaceTag($mapping,'#TAG507',number_abbr($SuspiciousEventsCount)); // Suspicious Domains ???? HIGH RISK ??
				$mapping = replaceTag($mapping,'#TAG508',number_abbr($PLACEHOLDER)); // First to Detect (Domain Count)
				$mapping = replaceTag($mapping,'#TAG509',number_abbr($PLACEHOLDER)); // Bandwidth Savings
		
				##// Slide 6 - Security Indicator Summary
				$mapping = replaceTag($mapping,'#TAG601',number_abbr($DNSActivityCount)); // DNS Requests
				$mapping = replaceTag($mapping,'#TAG602',number_abbr($HighEventsCount)); // High-Risk Events
				$mapping = replaceTag($mapping,'#TAG603',number_abbr($MediumEventsCount)); // Medium-Risk Events
				$mapping = replaceTag($mapping,'#TAG604',number_abbr($TotalInsights)); // Insights
				$mapping = replaceTag($mapping,'#TAG605',number_abbr($LookalikeThreatCount)); // Custom Lookalike Domains
				$mapping = replaceTag($mapping,'#TAG606',number_abbr($DOHEventsCount)); // DoH
				$mapping = replaceTag($mapping,'#TAG607',number_abbr($ZeroDayDNSEventsCount)); // Zero Day DNS

				// ** TODO ** //
				$mapping = replaceTag($mapping,'#TAG608',number_abbr($SuspiciousEventsCount)); // Suspicious Domains ???? HIGH RISK ??
		
				$mapping = replaceTag($mapping,'#TAG609',number_abbr($NODEventsCount)); // Newly Observed Domains
				$mapping = replaceTag($mapping,'#TAG610',number_abbr($DGAEventsCount)); // Domain Generated Algorithms
				$mapping = replaceTag($mapping,'#TAG611',number_abbr($DataExfilEventsCount)); // DNS Tunnelling
				$mapping = replaceTag($mapping,'#TAG612',number_abbr($UniqueApplicationsCount)); // Unique Applications
				$mapping = replaceTag($mapping,'#TAG613',number_abbr($HighRiskWebCategoryCount)); // High-Risk Web Categories
				$mapping = replaceTag($mapping,'#TAG614',number_abbr($ThreatActorsCountMetric)); // Threat Actors

				// ** TODO ** //
				$mapping = replaceTag($mapping,'#TAG615',number_abbr($PLACEHOLDER)); // First to Detect (Domain Count)
				$mapping = replaceTag($mapping,'#TAG616',number_abbr($PLACEHOLDER)); // Malicious TDS Events
				// This has been removed?
				$mapping = replaceTag($mapping,'#TAG22',number_abbr($DNSFirewallActivityDailyAverage)); // Avg Events Per Day
	
				##// Slide 8 - Additional Threat Intel Insights
				// ** CURRENTLY DONE MANUALLY BY ITI ** //
		
				##// Slide 11 - Traffic Usage Analysis
				// Total DNS Activity
				$mapping = replaceTag($mapping,'#TAG1101',number_abbr($DNSActivityCount));
				// DNS Firewall Activity
				$mapping = replaceTag($mapping,'#TAG1102',number_abbr($HML)); // Total
				$mapping = replaceTag($mapping,'#TAG1103',number_abbr($HighEventsCount)); // High Int
				$mapping = replaceTag($mapping,'#TAG1104',number_format($HighPerc,2).'%'); // High Percent
				$mapping = replaceTag($mapping,'#TAG1105',number_abbr($MediumEventsCount)); // Medium Int
				$mapping = replaceTag($mapping,'#TAG1106',number_format($MediumPerc,2).'%'); // Medium Percent
				$mapping = replaceTag($mapping,'#TAG1107',number_abbr($LowEventsCount)); // Low Int
				$mapping = replaceTag($mapping,'#TAG1108',number_format($LowPerc,2).'%'); // Low Percent
				// Threat Activity
				$mapping = replaceTag($mapping,'#TAG1109',number_abbr($ThreatActivityEventsCount));
				// Data Exfiltration Incidents
				$mapping = replaceTag($mapping,'#TAG1110',number_abbr($DataExfilEventsCount));
		
				##// Slide 12 - Traffic Analysis - DNS Activity
				$mapping = replaceTag($mapping,'#TAG1201',number_abbr($DNSActivityDailyAverage));
				##// Slide 13 - Traffic Analysis - DNS Firewall Activity
				$mapping = replaceTag($mapping,'#TAG1301',number_abbr($DNSFirewallActivityDailyAverage));
	
				##// Slide 15 - Key Insights
				// Insight Severity
				$mapping = replaceTag($mapping,'#TAG1501',number_abbr($TotalInsights)); // Total Open Insights
				$mapping = replaceTag($mapping,'#TAG1502',number_abbr($CriticalInsights)); // Critical Priority Insights
				$mapping = replaceTag($mapping,'#TAG1503',number_abbr($HighInsights)); // High Priority Insights
				$mapping = replaceTag($mapping,'#TAG1504',number_abbr($MediumInsights)); // Medium Priority Insights
				$mapping = replaceTag($mapping,'#TAG1505',number_abbr($LowInsights)); // Low Priority Insights
				$mapping = replaceTag($mapping,'#TAG1506',number_abbr($InfoInsights)); // Info Priority Insights
				// Event To Insight Aggregation
				$mapping = replaceTag($mapping,'#TAG1507',number_abbr($SecurityEventsCount)); // Events
				$mapping = replaceTag($mapping,'#TAG1508',number_abbr($TotalInsights)); // Key Insights
		
				##// Slide 24 - Lookalike Domains
				$mapping = replaceTag($mapping,'#TAG2501',number_abbr($LookalikeTotalCount)); // Total Lookalikes
				// if ($LookalikeTotalPercentage >= 0){$arrow='↑';} else {$arrow='↓';}
				// $mapping = replaceTag($mapping,'#TAG39',$arrow); // Arrow Up/Down
				// $mapping = replaceTag($mapping,'#TAG40',number_abbr($LookalikeTotalPercentage)); // Total Percentage Increase
				$mapping = replaceTag($mapping,'#TAG2502',number_abbr($LookalikeCustomCount)); // Total Lookalikes from Custom Watched Domains
				// if ($LookalikeCustomPercentage >= 0){$arrow='↑';} else {$arrow='↓';}
				// $mapping = replaceTag($mapping,'#TAG42',$arrow); // Arrow Up/Down
				// $mapping = replaceTag($mapping,'#TAG43',number_abbr($LookalikeCustomPercentage)); // Custom Percentage Increase
				$mapping = replaceTag($mapping,'#TAG2503',number_abbr($LookalikeThreatCount)); // Threats from Custom Watched Domains
				// if ($LookalikeThreatPercentage >= 0){$arrow='↑';} else {$arrow='↓';}
				// $mapping = replaceTag($mapping,'#TAG45',$arrow); // Arrow Up/Down
				// $mapping = replaceTag($mapping,'#TAG46',number_abbr($LookalikeThreatPercentage)); // Threats Percentage Increase
		
				##// Slide 29 - Security Activities
				$mapping = replaceTag($mapping,'#TAG2901',number_abbr($SecurityEventsCount)); // Security Events
				$mapping = replaceTag($mapping,'#TAG2902',number_abbr($DNSFirewallEventsCount)); // DNS Firewall
				$mapping = replaceTag($mapping,'#TAG2903',number_abbr($WebContentEventsCount)); // Web Content
				$mapping = replaceTag($mapping,'#TAG2904',number_abbr($DeviceCount)); // Devices
				$mapping = replaceTag($mapping,'#TAG2905',number_abbr($UserCount)); // Users
				$mapping = replaceTag($mapping,'#TAG2906',number_abbr($TotalInsights)); // Insights
				$mapping = replaceTag($mapping,'#TAG2907',number_abbr($ThreatInsightCount)); // Threat Insight
				$mapping = replaceTag($mapping,'#TAG2908',number_abbr($ThreatViewCount)); // Threat View
				$mapping = replaceTag($mapping,'#TAG2909',number_abbr($SourcesCount)); // Sources
		
				##// Slide 32 -> Onwards - Threat Actors
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
				'web_unique_applications' => $UniqueApplicationsCount,
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
			'Total' => ($Total * 13) + 27,
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