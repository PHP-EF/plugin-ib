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

	// public function TestExtract() {
	// 	extractZip($this->getDir()['Files'].'/templates/security-latest-(portrait).pptx',$this->getDir()['Files'].'/reports/test');
	// 	// extractZip($this->getDir()['Files'].'/reports/test1.pptx',$this->getDir()['Files'].'/reports/test');
	// }

	// public function TestManipulation() {
	// 	// Open PPTX Presentation _rels XML
	// 	$xml_rels = null;
	// 	$xml_rels = new DOMDocument('1.0', 'utf-8');
	// 	$xml_rels->formatOutput = true;
	// 	$xml_rels->preserveWhiteSpace = false;
	// 	$xml_rels->load($this->getDir()['Files'].'/reports/test/ppt/_rels/presentation.xml.rels');

	// 	// Open PPTX Presentation XML
	// 	$xml_pres = new DOMDocument('1.0', 'utf-8');
	// 	$xml_pres->formatOutput = true;
	// 	$xml_pres->preserveWhiteSpace = false;
	// 	$xml_pres->load($this->getDir()['Files'].'/reports/test/ppt/presentation.xml');
		
	// 	// Qty of slides to end up with
	// 	$SOCInsightsSlideCount = 2;

	// 	// New slides to be appended after this slide number
	// 	$SOCInsightsSlideStart = 17;
	// 	// Calculate the slide position based on above value
	// 	$SOCInsightsSlidePosition = $SOCInsightsSlideStart-2;
		
	// 	// Tag Numbers Start
	// 	$SITagStart = 100;
		
	// 	// Create Document Fragments for appending new relationships and set Starting Integers
	// 	$xml_rels_f = $xml_rels->createDocumentFragment();
	// 	$xml_rels_fstart = ($xml_rels->getElementsByTagName('Relationship')->length)+50;
	// 	$xml_pres_f = $xml_pres->createDocumentFragment();
	// 	$xml_pres_fstart = 13700;

	// 	// Get Slide Count
	// 	$SISlidesCount = iterator_count(new FilesystemIterator($this->getDir()['Files'].'/reports/test/ppt/slides'));
	// 	// Set first slide number
	// 	$SISlideNumber = $SISlidesCount++;

	// 	// Create new relationship / slide ID
	// 	$xml_rels_f->appendXML('<Relationship Id="rId150" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'.$SISlideNumber.'.xml"/>');
	// 	$xml_pres_f->appendXML('<p:sldId id="13700" r:id="rId150"/>');

	// 	// Clone Slides
	// 	copy($this->getDir()['Files'].'/reports/test/ppt/slides/slide'.$SOCInsightsSlideStart.'.xml',$this->getDir()['Files'].'/reports/test/ppt/slides/slide'.$SISlideNumber.'.xml');
	// 	copy($this->getDir()['Files'].'/reports/test/ppt/slides/_rels/slide'.$SOCInsightsSlideStart.'.xml.rels',$this->getDir()['Files'].'/reports/test/ppt/slides/_rels/slide'.$SISlideNumber.'.xml.rels');

	// 	// Append new relationship / slide ID
	// 	$xml_rels->getElementsByTagName('Relationships')->item(0)->appendChild($xml_rels_f);
	// 	$xml_pres->getElementsByTagName('sldId')->item($SOCInsightsSlidePosition)->after($xml_pres_f);

	// 	// Save core XML files
	// 	$xml_rels->save($this->getDir()['Files'].'/reports/test/ppt/_rels/presentation.xml.rels');
	// 	$xml_pres->save($this->getDir()['Files'].'/reports/test/ppt/presentation.xml');

	// 	// Load Slide XML _rels
	// 	$xml_sis = new DOMDocument('1.0', 'utf-8');
	// 	$xml_sis->formatOutput = true;
	// 	$xml_sis->preserveWhiteSpace = false;
	// 	$xml_sis->load($this->getDir()['Files'].'/reports/test/ppt/slides/_rels/slide'.$SISlideNumber.'.xml.rels');

	// 	// Remove notes
	// 	foreach ($xml_sis->getElementsByTagName('Relationship') as $element) {
	// 		// Remove notes references to avoid having to create unneccessary notes resources
	// 		if ($element->getAttribute('Type') == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide") {
	// 			$element->remove();
	// 		}
	// 		if ($element->getAttribute('Type') == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart") {
	// 			$OldChartNumber = str_replace('../charts/chart','',$element->getAttribute('Target'));
	// 			$OldChartNumber = str_replace('.xml','',$OldChartNumber);
	// 			$NewChartNumber = $OldChartNumber + 50;
	// 			copy($this->getDir()['Files'].'/reports/test/ppt/charts/chart'.$OldChartNumber.'.xml',$this->getDir()['Files'].'/reports/test/ppt/charts/chart'.$NewChartNumber.'.xml');
	// 			copy($this->getDir()['Files'].'/reports/test/ppt/charts/_rels/chart'.$OldChartNumber.'.xml.rels',$this->getDir()['Files'].'/reports/test/ppt/charts/_rels/chart'.$NewChartNumber.'.xml.rels');
	// 			$element->setAttribute('Target','../charts/chart'.$NewChartNumber.'.xml');

	// 			// Load Chart XML Rels
	// 			$xml_chart_rels = new DOMDocument('1.0', 'utf-8');
	// 			$xml_chart_rels->formatOutput = true;
	// 			$xml_chart_rels->preserveWhiteSpace = false;
	// 			$xml_chart_rels->load($this->getDir()['Files'].'/reports/test/ppt/charts/_rels/chart'.$NewChartNumber.'.xml.rels');

	// 			// Get colors and style relationships
	// 			// http://schemas.microsoft.com/office/2011/relationships/chartColorStyle
	// 			// http://schemas.microsoft.com/office/2011/relationships/chartStyle
	// 			foreach ($xml_chart_rels->getElementsByTagName('Relationship') as $element_c) {
	// 				if ($element_c->getAttribute('Type') == "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle") {
	// 					$OldColourNumber = str_replace('colors','',$element_c->getAttribute('Target'));
	// 					$OldColourNumber = str_replace('.xml','',$OldColourNumber);
	// 					$NewColourNumber = $OldColourNumber + 50;
	// 					copy($this->getDir()['Files'].'/reports/test/ppt/charts/colors'.$OldColourNumber.'.xml',$this->getDir()['Files'].'/reports/test/ppt/charts/colors'.$NewColourNumber.'.xml');
	// 					$element_c->setAttribute('Target','../charts/colors'.$NewColourNumber.'.xml');
	// 				} elseif ($element_c->getAttribute('Type') == "http://schemas.microsoft.com/office/2011/relationships/chartStyle") {
	// 					$OldStyleNumber = str_replace('style','',$element_c->getAttribute('Target'));
	// 					$OldStyleNumber = str_replace('.xml','',$OldStyleNumber);
	// 					$NewStyleNumber = $OldStyleNumber + 50;
	// 					copy($this->getDir()['Files'].'/reports/test/ppt/charts/style'.$OldStyleNumber.'.xml',$this->getDir()['Files'].'/reports/test/ppt/charts/style'.$NewStyleNumber.'.xml');
	// 					$element_c->setAttribute('Target','../charts/style'.$NewStyleNumber.'.xml');
	// 				} elseif ($element_c->getAttribute('Type') == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package") {
	// 					$OldEmbeddedNumber = str_replace('../embeddings/Microsoft_Excel_Worksheet','',$element_c->getAttribute('Target'));
	// 					$OldEmbeddedNumber = str_replace('.xlsx','',$OldEmbeddedNumber);
	// 					$NewEmbeddedNumber = $OldEmbeddedNumber + 50;
	// 					copy($this->getDir()['Files'].'/reports/test/ppt/embeddings/Microsoft_Excel_Worksheet'.$OldEmbeddedNumber.'.xlsx',$this->getDir()['Files'].'/reports/test/ppt/embeddings/Microsoft_Excel_Worksheet'.$NewEmbeddedNumber.'.xlsx');
	// 					$element_c->setAttribute('Target','../embeddings/Microsoft_Excel_Worksheet'.$NewEmbeddedNumber.'.xlsx');
	// 				}
	// 			}

	// 			$xml_chart_rels->save($this->getDir()['Files'].'/reports/test/ppt/charts/_rels/chart'.$NewChartNumber.'.xml.rels');
	// 		}
	// 	}

	// 	$xml_sis->save($this->getDir()['Files'].'/reports/test/ppt/slides/_rels/slide'.$SISlideNumber.'.xml.rels');
	// }

	// public function TestCompress() {
	// 	compressZip($this->getDir()['Files'].'/reports/test.pptx',$this->getDir()['Files'].'/reports/test');
	// }

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
				'MaliciousTDS' => '{"measures":["PortunusAggThreat_ch.requests","PortunusAggThreat_ch.timestampMax","PortunusAggThreat_ch.totalAssetCount","PortunusAggThreat_ch.threatCount"],"dimensions":["PortunusAggThreat_ch.severity"],"filters":[{"member":"PortunusAggThreat_ch.severity","operator":"notEquals","values":["Info"]},{"member":"PortunusAggThreat_ch.tclass","operator":"equals","values":["Malicious"]},{"member":"PortunusAggThreat_ch.tproperty","operator":"equals","values":["TDS"]}],"segments":["PortunusAggThreat_ch.threat_classes"],"timeDimensions":[{"dimension":"PortunusAggThreat_ch.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"limit":39,"offset":0,"total":true,"order":{"PortunusAggThreat_ch.timestampMax":"asc"}}',
				'IVThirtyDays' => '{"dimensions":["IVThirtyDays.dimensionLabel","IVThirtyDays.dimensionMetric","IVThirtyDays.industryName","IVThirtyDays.accountDimensionMetricSum","IVThirtyDays.industryDimensionMetricSum","IVThirtyDays.allAccountsDimensionMetricSum","IVThirtyDays.accountDimensionSum","IVThirtyDays.industryDimensionSum","IVThirtyDays.allAccountsDimensionSum","IVThirtyDays.accountPercentage","IVThirtyDays.industryPercentage","IVThirtyDays.overallPercentage"]}',
				'IVPeerAndOverallCount' => '{"measures":["IndustryVerticalPeerAndOverallCount.industryCount","IndustryVerticalPeerAndOverallCount.totalCount"]}',
				'IVThreatActors' => '{"dimensions":["IVThreatActors.accountThirtyDays","IVThreatActors.industryThirtyDays","IVThreatActors.allAccountThirtyDays"]}',
				'IVTrend' => '{"dimensions":["IVTrend.dimensionLabel","IVTrend.trendPercentage"]}',
				'SecurityTokens' => '{"ungrouped":false,"segments":[],"measures":["TokenUtilSecurityDaily.tokens","TokenUtilSecurityDaily.count"],"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null,"dimension":"TokenUtilSecurityDaily.timestamp"}],"dimensions":["TokenUtilSecurityDaily.account_id","TokenUtilSecurityDaily.category","TokenUtilSecurityDaily.type","TokenUtilSecurityDaily.timestamp","TokenUtilSecurityDaily.sub_type"]}',
				'ReportingTokens' => '{"ungrouped":false,"segments":[],"measures":["TokenUtilReportingDaily.tokens","TokenUtilReportingDaily.count"],"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null,"dimension":"TokenUtilReportingDaily.timestamp"}],"dimensions":["TokenUtilReportingDaily.account_id","TokenUtilReportingDaily.category","TokenUtilReportingDaily.timestamp"]}'
			);
			// Workaround for EU / US Realm Alignment
			// if ($config['Realm'] == 'EU') {
				// $CubeJSRequests['ThreatActors'] = '{"segments":[],"timeDimensions":[{"dimension":"PortunusAggIPSummary.timestamp","granularity":null,"dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"ungrouped":false,"order":{"PortunusAggIPSummary.timestampMax":"desc"},"measures":["PortunusAggIPSummary.count"],"dimensions":["PortunusAggIPSummary.threat_indicator","PortunusAggIPSummary.actor_id"],"limit":1000,"filters":[{"and":[{"operator":"set","member":"PortunusAggIPSummary.threat_indicator"},{"operator":"set","member":"PortunusAggIPSummary.actor_id"}]}]}';
			// } elseif ($config['Realm'] == 'US') {
				$CubeJSRequests['ThreatActors'] = '{"measures":[],"segments":[],"dimensions":["ThreatActors.storageid","ThreatActors.ikbactorid","ThreatActors.domain","ThreatActors.ikbfirstsubmittedts","ThreatActors.vtfirstdetectedts","ThreatActors.firstdetectedts","ThreatActors.lastdetectedts"],"timeDimensions":[{"dimension":"ThreatActors.lastdetectedts","granularity":null,"dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"ungrouped":false}';
			// }
			$CubeJSResults = $this->QueryCubeJSMulti($CubeJSRequests);

			// Define SOC Insights Cube JS Array
			$SOCInsightsCubeJSRequests = array();
	
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
				'BandwidthSavings' => 0,
				'TopDetectedProperties' => 5,
				'ContentFiltration' => 6,
				'DNSActivity' => 7,
				'DNSFirewallActivity' => 8,
				'InsightDistribution' => 9,
				// 10/11 - Outlier Insight
				'AppDiscovery' => 12,
				'WebContentDiscovery' => 13,
				'Lookalikes' => 14,
				'ZeroDayDNS' => array(
					'Landscape' => 15,
					'Portrait' => 15
				),
				'IVRiskyIndicatorsAccount' => array(
					'Landscape' => 16,
					'Portrait' => 16
				),
				'IVRiskyIndicatorsIndustry' => array(
					'Landscape' => 17,
					'Portrait' => 17
				),
				'IVThreatActorsAccount' => array(
					'Landscape' => 20,
					'Portrait' => 18
				),
				'IVThreatActorsIndustry' => array(
					'Landscape' => 21,
					'Portrait' => 19
				),
				'IVMaliciousIndicatorsAccount' => array(
					'Landscape' => 18,
					'Portrait' => 20
				),
				'IVMaliciousIndicatorsIndustry' => array(
					'Landscape' => 19,
					'Portrait' => 21
				),
				'IVThreatInsightAccount' => array(
					'Landscape' => 22,
					'Portrait' => 22
				),
				'IVThreatInsightIndustry' => array(
					'Landscape' => 23,
					'Portrait' => 23
				)
			];

			$CriticalalityMapping = [
				'5' => 'N/A',
				'4' => 'Critical',
				'3' => 'High',
				'2' => 'Medium',
				'1' => 'Low',
				'0' => 'Informational'
			];

			// 10 - Outlier Insight (Assets)
			// 11 - Outlier Insight (Indicators/Events)

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
			$TotalBandwidthSavedTop5Classes = array_slice($BandwidthSavings->result->data, 0, 5);

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
	
			// SOC Insights Details - Slide 16+
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building SOC Insight Details (This may take a moment)");
			$SOCInsightDetailsQuery = $this->queryCSP('get','/api/v1/insights?status=Active');
			$SOCInsightDetails = $SOCInsightDetailsQuery->insightList ?? [];
			// Create array for blocked/not blocked indicator counts by insightId
			$SOCInsightsIndicatorsBlockedCounts = [];
			
			// Build CubeJS Queries for each insight
			foreach ($SOCInsightDetails as $SID) {
				// General Info - Blocked Count, Confidence Level, Description, Feed Source, Insight Status, Insight ID, Most Recent At, Not Blocked Count, Persistent, Persistent Date, Spreading, Spreading Date, Started At, Threat Level, Threat Type, TClass, TFamily
				$SOCInsightsCubeJSRequests[$SID->insightId.'-general'] = '{"order":{"InsightDetails.mostRecentAt":"desc"},"timeDimensions":[{"dimension":"InsightDetails.eventSummaryHour","dateRange":["'.$StartDimension.'","'.$EndDimension.'"]}],"measures":["InsightDetails.blockedCount","InsightDetails.notBlockedCount","InsightDetails.numEvents","InsightDetails.mostRecentAt"],"dimensions":["InsightDetails.insightId","InsightDetails.tClass","InsightDetails.tFamily","InsightDetails.threatType","InsightDetails.description","InsightDetails.threatLevel","InsightDetails.feedSource","InsightDetails.startedAt","InsightDetails.confidenceLevel","InsightDetails.insightStatus","InsightDetails.persistent","InsightDetails.persistentDate","InsightDetails.spreading","InsightDetails.spreadingDate"],"filters":[{"member":"InsightDetails.insightId","operator":"equals","values":["'.$SID->insightId.'"]}],"timezone":"UTC"}';

				// Total Asset Count / Verified Asset Count / Unverified Asset Count
				$SOCInsightsCubeJSRequests[$SID->insightId.'-assetCounts'] = '{"filters":[{"member":"PortunusAggIPSummary_ch.tclass","values":["'.$SID->tClass.'"],"operator":"equals"},{"member":"PortunusAggIPSummary_ch.tfamily","values":["'.$SID->tFamily.'"],"operator":"equals"}],"segments":[],"dimensions":["PortunusAggIPSummary_ch.tfamily","PortunusAggIPSummary_ch.tclass"],"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"dimension":"PortunusAggIPSummary_ch.timestamp","granularity":null}],"measures":["PortunusAggIPSummary_ch.unknownAssetCount","PortunusAggIPSummary_ch.knownAssetCount","PortunusAggIPSummary_ch.totalAssetCount"],"ungrouped":false}';

				// Get Total Number of Unique Assets Accessing Unblocked Indicators (deviceId Distinct Count)
				$SOCInsightsCubeJSRequests[$SID->insightId.'-assetsAccessingDeviceID'] = '{"order":{"PortunusAggIPSummary.timestamp":"desc"},"measures":["PortunusAggIPSummary.deviceIdDistinctCount"],"timeDimensions":[{"dimension":"PortunusAggIPSummary.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null}],"filters":[{"member":"PortunusAggIPSummary.tclass","operator":"equals","values":["'.$SID->tClass.'"]},{"member":"PortunusAggIPSummary.tfamily","operator":"equals","values":["'.$SID->tFamily.'"]},{"member":"PortunusAggIPSummary.device_id","operator":"set"},{"member":"PortunusAggIPSummary.action","operator":"equals","values":["Not Blocked"]}]}';

				// Get Total Number of Unique Assets Accessing Unblocked Indicators (qip Distinct Count)
				$SOCInsightsCubeJSRequests[$SID->insightId.'-assetsAccessingQIP'] = '{"order":{"PortunusAggIPSummary.timestamp":"desc"},"measures":["PortunusAggIPSummary.qipDistinctCount"],"timeDimensions":[{"dimension":"PortunusAggIPSummary.timestamp","dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"granularity":null}],"filters":[{"member":"PortunusAggIPSummary.tclass","operator":"equals","values":["'.$SID->tClass.'"]},{"member":"PortunusAggIPSummary.tfamily","operator":"equals","values":["'.$SID->tFamily.'"]},{"member":"PortunusAggIPSummary.device_id","operator":"notSet"},{"member":"PortunusAggIPSummary.action","operator":"equals","values":["Not Blocked"]}]}';

				// Get Total Blocked/Not Blocked Indicator Counts
				$IndicatorStart = (new DateTime($StartDimension))->format('Y-m-d\TH:i:s').'.000';
				$IndicatorEnd = (new DateTime($EndDimension))->format('Y-m-d\TH:i:s').'.000';
				// Should this be insight_type=detections or insight_type=rpz ? How do I determine this?
				$SOCInsightsIndicatorsBlockedCounts[$SID->insightId] = $this->queryCSP("get","api/ris/v1/insights/indicators/counts?tclass=".$SID->tClass."&tfamily=".$SID->tFamily."&insight_type=detections&from=".$IndicatorStart."&to=".$IndicatorEnd);
				$SOCInsightsIndicatorsDetails[$SID->insightId] = $this->queryCSP("get","api/ris/v1/insights/indicators/details?tclass=".$SID->tClass."&tfamily=".$SID->tFamily."&insight_type=detections&from=".$IndicatorStart."&to=".$IndicatorEnd."&limit=1000");

				// Get Impacted Assets & Indicators Time Series
				$SOCInsightsCubeJSRequests[$SID->insightId.'-impactedAssetsTimeSeries'] = '{"order":{"PortunusAggIPSummary.timestamp":"desc"},"measures":["PortunusAggIPSummary.deviceIdDistinctCount"],"dimensions":[],"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"dimension":"PortunusAggIPSummary.timestamp","granularity":"hour"}],"filters":[{"member":"PortunusAggIPSummary.tclass","operator":"equals","values":["'.$SID->tClass.'"]},{"member":"PortunusAggIPSummary.tfamily","operator":"equals","values":["'.$SID->tFamily.'"]},{"member":"PortunusAggIPSummary.device_id","operator":"set"}]}';
				$SOCInsightsCubeJSRequests[$SID->insightId.'-indicatorsTimeSeries'] = '{"timeDimensions":[{"dateRange":["'.$StartDimension.'","'.$EndDimension.'"],"dimension":"PortunusAggIPSummary.timestamp","granularity":"hour"}],"measures":["PortunusAggIPSummary.threatIndicatorDistinctCount"],"dimensions":["PortunusAggIPSummary.threat_indicator"],"filters":[{"member":"PortunusAggIPSummary.tclass","operator":"equals","values":["'.$SID->tClass.'"]},{"member":"PortunusAggIPSummary.tfamily","operator":"equals","values":["'.$SID->tFamily.'"]}]}';
			}
			// Invoke CubeJS to populate SOC Insight Data
			$SOCInsightsCubeJSResults = $this->QueryCubeJSMulti($SOCInsightsCubeJSRequests);

			// ** Industry Vertical Start 
			// ** //
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Industry Vertical Metrics");
			$IVThirtyDays = $CubeJSResults['IVThirtyDays']['Body'];
			$IVPeerAndOverallCount = $CubeJSResults['IVPeerAndOverallCount']['Body'];
			$IVThreatActors = $CubeJSResults['IVThreatActors']['Body'];
			$IVTrend = $CubeJSResults['IVTrend']['Body'];


			// ** INDUSTRY VERTICAL NAME & PEER COUNT ** //
			$IVIndustryName = $IVThirtyDays->result->data[0]->{'IVThirtyDays.industryName'} ?? 'Unknown';
			$IVIndustryPeerCount = $IVPeerAndOverallCount->result->data[0]->{'IndustryVerticalPeerAndOverallCount.industryCount'} ?? 0;

			// ** CONFIRMED THREATS SEEN START ** //

			// Find each object under $IVThirtyDays->result->data with dimensionLabel of X
			$DLConfirmedThreatsSeen = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'confirmed_threats_seen';
			});
			$DLConfirmedThreatsSeen = array_values($DLConfirmedThreatsSeen);

			// Determine Customer Totals/Averages & Trend
			$IVPDLConfirmedThreatsSeenSum = $DLConfirmedThreatsSeen[0]->{'IVThirtyDays.accountDimensionSum'} ?? 0;
			$IVPDLConfirmedThreatsSeenPercentage = $DLConfirmedThreatsSeen[0]->{'IVThirtyDays.accountPercentage'} ?? 0;
			$IVPDLConfirmedThreatsSeenTrend = $IVTrend->result->data[array_search('confirmed_threats_seen', array_column($IVTrend->result->data, 'IVTrend.dimensionLabel'))]->{'IVTrend.trendPercentage'} ?? 0;

			// Determine Total Number of Malicious Domains vs Malicious IPs
			$IVPDLMaliciousDomains = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'confirmed_threats_seen' && $item->{'IVThirtyDays.dimensionMetric'} === 'infoblox-base';
			});
			$IVPDLMaliciousDomains = array_values($IVPDLMaliciousDomains);
			$IVPDLMaliciousIPs = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'confirmed_threats_seen' && $item->{'IVThirtyDays.dimensionMetric'} === 'infoblox-base-ip';
			});
			$IVPDLMaliciousIPs = array_values($IVPDLMaliciousIPs);
			// Work out percentage of Malicious Domains vs.	 Malicious IPs (Account)
			if ($IVPDLConfirmedThreatsSeenSum > 0) {
				$IVPDLMaliciousDomainsPercentage = round((($IVPDLMaliciousDomains[0]->{'IVThirtyDays.accountDimensionMetricSum'} ?? 0) / $IVPDLConfirmedThreatsSeenSum) * 100, 2);
				$IVPDLMaliciousIPsPercentage = round((($IVPDLMaliciousIPs[0]->{'IVThirtyDays.accountDimensionMetricSum'} ?? 0) / $IVPDLConfirmedThreatsSeenSum) * 100, 2);
			} else {
				$IVPDLMaliciousDomainsPercentage = 0;
				$IVPDLMaliciousIPsPercentage = 0;
			}

			// Determine Industry Totals/Averages
			$IVIDLConfirmedThreatsSeenSum = $DLConfirmedThreatsSeen[0]->{'IVThirtyDays.industryDimensionSum'} ?? 0;
			// Find IndustryVerticalPeerAndOverallCount.industryCount
			$IVIDLConfirmedThreatsSeenPercentage = $DLConfirmedThreatsSeen[0]->{'IVThirtyDays.industryPercentage'} ?? 0;
			// Work out percentage of Malicious Domains vs.	 Malicious IPs (Industry)
			$IVIDLTotalMaliciousIndicators = ($DLConfirmedThreatsSeen[0]->{'IVThirtyDays.industryDimensionMetricSum'} ?? 0);
			if ($IVIDLTotalMaliciousIndicators > 0) {
				$IVIDLMaliciousDomainsPercentage = round((($IVPDLMaliciousDomains[0]->{'IVThirtyDays.industryDimensionMetricSum'} ?? 0) / $IVIDLTotalMaliciousIndicators) * 100, 2);
				$IVIDLMaliciousIPsPercentage = round((($IVPDLMaliciousIPs[0]->{'IVThirtyDays.industryDimensionMetricSum'} ?? 0) / $IVIDLTotalMaliciousIndicators) * 100, 2);
			} else {
				$IVIDLMaliciousDomainsPercentage = 0;
				$IVIDLMaliciousIPsPercentage = 0;
			}

			// ** CONFIRMED THREATS SEEN END ** //


			// ** UNCONFIRMED THREATS SEEN START ** //
			$DLUnconfirmedThreatsSeen = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'unconfirmed_threats_seen';
			});
			$DLUnconfirmedThreatsSeen = array_values($DLUnconfirmedThreatsSeen);
			// Determine Customer Totals/Averages & Trend
			$IVPDLUnconfirmedThreatsSeenSum = $DLUnconfirmedThreatsSeen[0]->{'IVThirtyDays.accountDimensionSum'} ?? 0;
			$IVPDLUnconfirmedThreatsSeenPercentage = $DLUnconfirmedThreatsSeen[0]->{'IVThirtyDays.accountPercentage'} ?? 0;
			$IVPDLUnconfirmedThreatsSeenTrend = $IVTrend->result->data[array_search('unconfirmed_threats_seen', array_column($IVTrend->result->data, 'IVTrend.dimensionLabel'))]->{'IVTrend.trendPercentage'} ?? 0;

			// Determine Total Number of High Risk, Med Risk & Low Risk indicators, ensuring it is the first key in the array (key 0)
			$IVPDLUnconfirmedHighRisk = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'unconfirmed_threats_seen' && $item->{'IVThirtyDays.dimensionMetric'} === 'infoblox-high-risk';
			});
			$IVPDLUnconfirmedHighRisk = array_values($IVPDLUnconfirmedHighRisk);
			$IVPDLUnconfirmedMedRisk = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'unconfirmed_threats_seen' && $item->{'IVThirtyDays.dimensionMetric'} === 'infoblox-med-risk';
			});
			$IVPDLUnconfirmedMedRisk = array_values($IVPDLUnconfirmedMedRisk);
			$IVPDLUnconfirmedLowRisk = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'unconfirmed_threats_seen' && $item->{'IVThirtyDays.dimensionMetric'} === 'infoblox-low-risk';
			});
			$IVPDLUnconfirmedLowRisk = array_values($IVPDLUnconfirmedLowRisk);

			// Work out percentage of High Risk, Med Risk & Low Risk indicators (Account)
			if ($IVPDLUnconfirmedThreatsSeenSum > 0) {
				$IVPDLUnconfirmedHighRiskPercentage = round((($IVPDLUnconfirmedHighRisk[0]->{'IVThirtyDays.accountDimensionMetricSum'} ?? 0) / $IVPDLUnconfirmedThreatsSeenSum) * 100, 2);
				$IVPDLUnconfirmedMedRiskPercentage = round((($IVPDLUnconfirmedMedRisk[0]->{'IVThirtyDays.accountDimensionMetricSum'} ?? 0) / $IVPDLUnconfirmedThreatsSeenSum) * 100, 2);
				$IVPDLUnconfirmedLowRiskPercentage = round((($IVPDLUnconfirmedLowRisk[0]->{'IVThirtyDays.accountDimensionMetricSum'} ?? 0) / $IVPDLUnconfirmedThreatsSeenSum) * 100, 2);
			} else {
				$IVPDLUnconfirmedHighRiskPercentage = 0;
				$IVPDLUnconfirmedMedRiskPercentage = 0;
				$IVPDLUnconfirmedLowRiskPercentage = 0;
			}

			// Determine Industry Totals/Averages
			$IVIDLUnconfirmedThreatsSeenSum = $DLUnconfirmedThreatsSeen[0]->{'IVThirtyDays.industryDimensionSum'} ?? 0;
			$IVIDLUnconfirmedThreatsSeenPercentage = $DLUnconfirmedThreatsSeen[0]->{'IVThirtyDays.industryPercentage'} ?? 0;

			// Work out percentage of High Risk, Med Risk & Low Risk indicators (Industry)
			if ($IVIDLUnconfirmedThreatsSeenSum > 0) {
				$IVIDLUnconfirmedHighRiskPercentage = round((($IVPDLUnconfirmedHighRisk[0]->{'IVThirtyDays.industryDimensionMetricSum'} ?? 0) / $IVIDLUnconfirmedThreatsSeenSum) * 100, 2);
				$IVIDLUnconfirmedMedRiskPercentage = round((($IVPDLUnconfirmedMedRisk[0]->{'IVThirtyDays.industryDimensionMetricSum'} ?? 0) / $IVIDLUnconfirmedThreatsSeenSum) * 100, 2);
				$IVIDLUnconfirmedLowRiskPercentage = round((($IVPDLUnconfirmedLowRisk[0]->{'IVThirtyDays.industryDimensionMetricSum'} ?? 0) / $IVIDLUnconfirmedThreatsSeenSum) * 100, 2);
			} else {
				$IVIDLUnconfirmedHighRiskPercentage = 0;
				$IVIDLUnconfirmedMedRiskPercentage = 0;
				$IVIDLUnconfirmedLowRiskPercentage = 0;
			}
			// ** UNCONFIRMED THREATS SEEN END ** //


			// ** THREAT ACTOR ASSOCIATED TRAFFIC SEEN START ** //
			$DLThreatActorAssociatedTrafficSeen = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'threat_actor_associated_traffic_seen';
			});
			$DLThreatActorAssociatedTrafficSeen = array_values($DLThreatActorAssociatedTrafficSeen);
			// Determine Customer Totals/Averages & Trend
			$IVPDLThreatActorAssociatedTrafficSeenSum = $DLThreatActorAssociatedTrafficSeen[0]->{'IVThirtyDays.accountDimensionSum'} ?? 0;
			$IVPDLThreatActorAssociatedTrafficSeenPercentage = $DLThreatActorAssociatedTrafficSeen[0]->{'IVThirtyDays.accountPercentage'} ?? 0;
			$IVPDLThreatActorAssociatedTrafficSeenTrend = $IVTrend->result->data[array_search('threat_actor_associated_traffic_seen', array_column($IVTrend->result->data, 'IVTrend.dimensionLabel'))]->{'IVTrend.trendPercentage'} ?? 0;

			// Determine Industry Totals/Averages
			$IVIDLThreatActorAssociatedTrafficSeenSum = $DLThreatActorAssociatedTrafficSeen[0]->{'IVThirtyDays.industryDimensionSum'} ?? 0;
			$IVIDLThreatActorAssociatedTrafficSeenPercentage = $DLThreatActorAssociatedTrafficSeen[0]->{'IVThirtyDays.industryPercentage'} ?? 0;

			// Identify number of Threat Actors from $IVThreatActors
			if (isset($IVThreatActors->result->data[0])) {
				$IVThreatActorsMetrics = [
					'account' => $IVThreatActors->result->data[0]->{'IVThreatActors.accountThirtyDays'} ?? 0,
					'industry' => ($IVThreatActors->result->data[0]->{'IVThreatActors.industryThirtyDays'} ?? 0) / ($IVIndustryPeerCount > 0 ? $IVIndustryPeerCount : 1),
					'all' => $IVThreatActors->result->data[0]->{'IVThreatActors.allAccountThirtyDays'} ?? 0
				];
			} else {
				$IVThreatActorsMetrics = [
					'accounts' => 0,
					'industry' => 0,
					'all' => 0
				];
			}

			// ** THREAT ACTOR ASSOCIATED TRAFFIC SEEN END ** //

			// ** ZERO DAY DNS TRAFFIC SEEN START ** //
			$DLZeroDayDNSTrafficSeen = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'zero_day_dns_traffic_seen';
			});
			$DLZeroDayDNSTrafficSeen = array_values($DLZeroDayDNSTrafficSeen);

			// Determine Customer Totals/Averages & Trend
			$IVPDLZeroDayDNSTrafficSeenSum = $DLZeroDayDNSTrafficSeen[0]->{'IVThirtyDays.accountDimensionSum'} ?? 0;
			$IVPDLZeroDayDNSTrafficSeenPercentage = $DLZeroDayDNSTrafficSeen[0]->{'IVThirtyDays.accountPercentage'} ?? 0;
			$IVPDLZeroDayDNSTrafficSeenTrend = $IVTrend->result->data[array_search('zero_day_dns_traffic_seen', array_column($IVTrend->result->data, 'IVTrend.dimensionLabel'))]->{'IVTrend.trendPercentage'} ?? 0;

			// Determine Industry Totals/Averages
			$IVIDLZeroDayDNSTrafficSeenSum = $DLZeroDayDNSTrafficSeen[0]->{'IVThirtyDays.industryDimensionSum'} ?? 0;
			$IVIDLZeroDayDNSTrafficSeenPercentage = $DLZeroDayDNSTrafficSeen[0]->{'IVThirtyDays.industryPercentage'} ?? 0;

			// ** ZERO DAY DNS TRAFFIC SEEN END ** //

			// ** THREAT INSIGHT DETECTION START ** //
			$DLThreatInsightDetection = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'threat_insight_detection';
			});
			$DLThreatInsightDetection = array_values($DLThreatInsightDetection);

			// Determine Customer Totals/Averages & Trend
			$IVPDLThreatInsightDetectionSum = $DLThreatInsightDetection[0]->{'IVThirtyDays.accountDimensionSum'} ?? 0;
			$IVPDLThreatInsightDetectionPercentage = $DLThreatInsightDetection[0]->{'IVThirtyDays.accountPercentage'} ?? 0;
			$IVPDLThreatInsightDetectionTrend = $IVTrend->result->data[array_search('threat_insight_detection', array_column($IVTrend->result->data, 'IVTrend.dimensionLabel'))]->{'IVTrend.trendPercentage'} ?? 0;

			// Determine Industry Totals/Averages
			$IVIDLThreatInsightDetectionSum = $DLThreatInsightDetection[0]->{'IVThirtyDays.industryDimensionSum'} ?? 0;
			$IVIDLThreatInsightDetectionPercentage = $DLThreatInsightDetection[0]->{'IVThirtyDays.industryPercentage'} ?? 0;

			// Determine Total Number of DNST vs. DGA traffic, ensuring it is the first key in the array (key 0)
			$IVPDLDNSTraffic = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'threat_insight_detection' && $item->{'IVThirtyDays.dimensionMetric'} === 'dnst';
			});
			$IVPDLDNSTraffic = array_values($IVPDLDNSTraffic);

			$IVPDLDGATraffic = array_filter($IVThirtyDays->result->data, function($item) {
				return $item->{'IVThirtyDays.dimensionLabel'} === 'threat_insight_detection' && $item->{'IVThirtyDays.dimensionMetric'} === 'dga';
			});
			$IVPDLDGATraffic = array_values($IVPDLDGATraffic);

			// Work out the percentage of DNST vs. DGA traffic (Account)
			if ($IVPDLThreatInsightDetectionSum > 0) {
				$IVPDLDNSTTrafficPercentage = round((($IVPDLDNSTraffic[0]->{'IVThirtyDays.accountDimensionMetricSum'} ?? 0) / $IVPDLThreatInsightDetectionSum) * 100, 2);
				$IVPDLDGATrafficPercentage = round((($IVPDLDGATraffic[0]->{'IVThirtyDays.accountDimensionMetricSum'} ?? 0) / $IVPDLThreatInsightDetectionSum) * 100, 2);
			} else {
				$IVPDLDNSTTrafficPercentage = 0;
				$IVPDLDGATrafficPercentage = 0;
			}

			// Work out the percentage of DNST vs. DGA traffic (Industry)
			$IVIDLTotalThreatInsightDetection = ($DLThreatInsightDetection[0]->{'IVThirtyDays.industryDimensionMetricSum'} ?? 0);
			if ($IVIDLThreatInsightDetectionSum > 0) {
				$IVIDLDNSTTrafficPercentage = round((($IVPDLDNSTraffic[0]->{'IVThirtyDays.industryDimensionMetricSum'} ?? 0) / $IVIDLThreatInsightDetectionSum) * 100, 2);
				$IVIDLDGATrafficPercentage = round((($IVPDLDGATraffic[0]->{'IVThirtyDays.industryDimensionMetricSum'} ?? 0) / $IVIDLThreatInsightDetectionSum) * 100, 2);
			} else {
				$IVIDLDNSTTrafficPercentage = 0;
				$IVIDLDGATrafficPercentage = 0;
			}

			// ** THREAT INSIGHT DETECTION END ** //

			// Optional Query - Not currently used
			// $DLTotalDNSTraffic = array_filter($IVThirtyDays->result->data, function($item) {
			// 	return $item->{'IVThirtyDays.dimensionLabel'} === 'total_dns_traffic';
			// });

			// ** Industry Vertical End ** //

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
			} elseif ($DateDiff > 30) {
				$ZDDStartDimension = (new DateTime($EndDimension))->modify('-30 days')->format('Y-m-d\T').'00:00:00';
			} else {
				$ZDDStartDimension = $StartDimension;
			}
			// Fix some weird behaviour with time
			$ZDDEndDimension = (new DateTime($EndDimension))->modify('+1 hour')->format('Y-m-d\TH:i:s');
			$ZeroDayDNSDetectionsUri = 'tide-rpz-stats/v1/zero_day_detections?from='.$ZDDStartDimension.'Z&to='.$ZDDEndDimension.'Z';
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

			// Token Calculator
			$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Token Usage");
			$SecurityTokenCategories = [
				'SOC Insights' => [
					'Verified Assets' => 0,
					'Unverified Assets' => 0,
					'Total'	=> 0
				],
				'Lookalikes' => [
					'Domains' => 0,
					'Total'	=> 0
				],
				'Cloud Asset Protection' => [
					'Verified Assets' => 0,
					'Unverified Assets' => 0,
					'Total'	=> 0
				],
				'Dossier' => [
					'Queries' => 0,
					'Total'	=> 0
				],
				'Threat Defense for NIOS' => [
					'X5' => 0,
					'X6' => 0,
					'Total'	=> 0
				],
				'Total' => 0
			];
			$ReportingTokenCategories = [
				'30-day Active Search' => 0,
				'Ecosystem'	=> 0,
				'S3' => 0,
				'Total' => 0
			];

			$SecurityTokenUsage = $CubeJSResults['SecurityTokens']['Body']->result->data ?? [];
			$SecurityTokenDailyUsage = [];
			if (isset($SecurityTokenUsage)) {
				foreach ($SecurityTokenUsage as $STU) {
					// Group categories by timestamp, then identify the date with peak usage. The peak usage should then be used to populate $SecurityTokenCategories, including the respective totals
					$timestamp = $STU->{'TokenUtilSecurityDaily.timestamp'};
					$date = date('Y-m-d', strtotime($timestamp));
					if (!isset($SecurityTokenDailyUsage[$date])) {
						$SecurityTokenDailyUsage[$date] = [];
					}
					$SecurityTokenDailyUsage[$date][] = $STU;

				}
				// Now we have an array of daily usage, identify the peak day
				$SecurityPeakDate = null;
				$SecurityPeakTotal = 0;
				foreach ($SecurityTokenDailyUsage as $date => $usages) {
					$DailyTotal = array_sum(array_column($usages, 'TokenUtilSecurityDaily.tokens'));
					if ($DailyTotal > $SecurityPeakTotal) {
						$SecurityPeakTotal = $DailyTotal;
						$SecurityPeakDate = $date;
					}
				}
				// Now populate $SecurityTokenCategories based on the peak day usages
				if ($SecurityPeakDate !== null) {
					foreach ($SecurityTokenDailyUsage[$SecurityPeakDate] as $STU) {
						$category = $STU->{'TokenUtilSecurityDaily.category'};
						$type = $STU->{'TokenUtilSecurityDaily.type'};
						$count = $STU->{'TokenUtilSecurityDaily.tokens'};
						if (isset($SecurityTokenCategories[$category])) {
							if (in_array($category, ['SOC Insights', 'Cloud Asset Protection']) && in_array($type, ['Verified Assets', 'Unverified Assets'])) {
								$SecurityTokenCategories[$category][$type] += $count;
								$SecurityTokenCategories[$category]['Total'] += $count;
							} elseif ($category === 'Lookalikes' && $type === 'Domains') {
								$SecurityTokenCategories[$category]['Domains'] += $count;
								$SecurityTokenCategories[$category]['Total'] += $count;
							} elseif ($category === 'Dossier' && $type === 'Queries') {
								$SecurityTokenCategories[$category]['Queries'] += $count;
								$SecurityTokenCategories[$category]['Total'] += $count;
							} elseif ($category === 'Threat Defense for NIOS' && in_array($type, ['X5', 'X6'])) {
								$SecurityTokenCategories[$category][$type] += $count;
								$SecurityTokenCategories[$category]['Total'] += $count;
							}
							$SecurityTokenCategories['Total'] += $count;
						}
					}
				}
			}

			$ReportingTokenUsage = $CubeJSResults['ReportingTokens']['Body']->result->data ?? [];
			$ReportingTokenDailyUsage = [];
			if (isset($ReportingTokenUsage)) {
				foreach ($ReportingTokenUsage as $RTU) {
					// Group categories by timestamp, then identify the date with peak usage. The peak usage should then be used to populate $ReportingTokenCategories, including the respective totals
					$timestamp = $RTU->{'TokenUtilReportingDaily.timestamp'};
					$date = date('Y-m-d', strtotime($timestamp));
					if (!isset($ReportingTokenDailyUsage[$date])) {
						$ReportingTokenDailyUsage[$date] = [];
					}
					$ReportingTokenDailyUsage[$date][] = $RTU;
				}
				// Now we have an array of daily usage, identify the peak day
				$ReportingPeakDate = null;
				$ReportingPeakTotal = 0;
				foreach ($ReportingTokenDailyUsage as $date => $usages) {
					$DailyTotal = array_sum(array_column($usages, 'TokenUtilReportingDaily.tokens'));
					if ($DailyTotal > $ReportingPeakTotal) {
						$ReportingPeakTotal = $DailyTotal;
						$ReportingPeakDate = $date;
					}
				}
				// Now populate $ReportingTokenCategories based on the peak day usages. If the category is '30-day Active Search', exclude it from the total
				if ($ReportingPeakDate !== null) {
					foreach ($ReportingTokenDailyUsage[$ReportingPeakDate] as $RTU) {
						$category = $RTU->{'TokenUtilReportingDaily.category'};
						$count = $RTU->{'TokenUtilReportingDaily.tokens'};
						if (isset($ReportingTokenCategories[$category])) {
							$ReportingTokenCategories[$category] += $count;
							if ($category !== '30-day Active Search') {
								$ReportingTokenCategories['Total'] += $count;
							}
						}
					}
				}
			}
			// End of Token Calculator

			// New Loop for support of multiple selected templates
			// writeProgress result should be Selected Template 27 + 13 x N
			foreach ($SelectedTemplates as &$SelectedTemplate) {
				$embeddedDirectory = $SelectedTemplate['ExtractedDir'].'/ppt/embeddings/';
				$embeddedFiles = array_values(array_diff(scandir($embeddedDirectory), array('.', '..')));
				usort($embeddedFiles, 'strnatcmp');

				// Initialise array to hold references to excel files created for charts
				$SOCInsightsExcelReference = [];

				$this->logging->writeLog("Assessment","Embedded Files List","debug",['Template' => $SelectedTemplate, 'Embedded Files' => $embeddedFiles]);
	
				// Top detected properties
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Threat Properties");
				$TopDetectedProperties = $CubeJSResults['TopDetectedProperties']['Body'];
				if (isset($TopDetectedProperties->result->data)) {
					$EmbeddedTopDetectedProperties = getEmbeddedSheetFilePath('TopDetectedProperties', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
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
					$EmbeddedContentFiltration = getEmbeddedSheetFilePath('ContentFiltration', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
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
					$EmbeddedDNSActivityDaily = getEmbeddedSheetFilePath('DNSActivity', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
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
					$EmbeddedDNSFirewallActivityDaily = getEmbeddedSheetFilePath('DNSFirewallActivity', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
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
					$EmbeddedInsightDistribution = getEmbeddedSheetFilePath('InsightDistribution', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
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
					$EmbeddedLookalikes = getEmbeddedSheetFilePath('Lookalikes', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
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

				// Bandwidth Savings - Slide 7
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Bandwidth Savings Page");
				$EmbeddedBandwidthSavings = getEmbeddedSheetFilePath('BandwidthSavings', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$BandwidthSavingsSS = IOFactory::load($EmbeddedBandwidthSavings);
				$RowNo = 2;
				if (isset($TotalBandwidthSavedTop5Classes) AND is_array($TotalBandwidthSavedTop5Classes)) {
					foreach ($TotalBandwidthSavedTop5Classes as $BandwidthClass) {
						$BandwidthSavingsS = $BandwidthSavingsSS->getActiveSheet();
						$BandwidthSavingsS->setCellValue('A'.$RowNo, $BandwidthClass->{'PortunusAggThreat_ch.tclass'});
						// Work out percentage from $TotalBandwidthBytes
						if ($TotalBandwidthBytes > 0) {
							$BandwidthClassPercentage = ($BandwidthClass->{'PortunusAggThreat_ch.bandwidthTotal'} / $TotalBandwidthBytes) * 100;
						} else {
							$BandwidthClassPercentage = 0;
						}
						$BandwidthSavingsS->setCellValue('B'.$RowNo, $BandwidthClassPercentage);
						$RowNo++;
					}
				}
				$BandwidthSavingsW = IOFactory::createWriter($BandwidthSavingsSS, 'Xlsx');
				$BandwidthSavingsW->save($EmbeddedBandwidthSavings);

				// Application Detection - Slide 19
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Application Detection Page");
				$AppDiscoveryApplications->result->data = array_slice($AppDiscoveryApplications->result->data, 0, 10);
				$EmbeddedAppDiscovery = getEmbeddedSheetFilePath('AppDiscovery', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
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
				$EmbeddedWebContentDiscovery = getEmbeddedSheetFilePath('WebContentDiscovery', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
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

				// Zero Day DNS - Slide 35 / 38
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Zero Day DNS Page");
				$EmbeddedZeroDayDNS = getEmbeddedSheetFilePath('ZeroDayDNS', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$ZeroDayDNSSS = IOFactory::load($EmbeddedZeroDayDNS);
				// Get first 5 detections from $ZeroDayDNSDetections->detections
				$RowNo = 2;
				if (isset($ZeroDayDNSDetections->detections) AND is_array($ZeroDayDNSDetections->detections)) {
					$ZeroDayDNSDetectionsSliced = array_slice($ZeroDayDNSDetections->detections, 0, 5);
					foreach ($ZeroDayDNSDetectionsSliced as $ZeroDayDNSDetection) {
						$ZeroDayDNSS = $ZeroDayDNSSS->getActiveSheet();
						$ZeroDayDNSS->setCellValue('A'.$RowNo, $ZeroDayDNSDetection->domain ?? 'Unknown');
						$ZeroDayDNSS->setCellValue('B'.$RowNo, $ZeroDayDNSDetection->count ?? 0);
						$RowNo++;
					}
				}
				$ZeroDayDNSW = IOFactory::createWriter($ZeroDayDNSSS, 'Xlsx');
				$ZeroDayDNSW->save($EmbeddedZeroDayDNS);

				// ** Industry Vertical (Slide 37/39 On Template) ** //
				$Progress = $this->writeProgress($config['UUID'],$Progress,"Building Industry Vertical Page");

				// Malicious Indicators Seen - Account
				$EmbeddedIVMaliciousIndicatorsAccount = getEmbeddedSheetFilePath('IVMaliciousIndicatorsAccount', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$IVMaliciousIndicatorsAccountSS = IOFactory::load($EmbeddedIVMaliciousIndicatorsAccount);
				$IVMaliciousIndicatorsAccountS = $IVMaliciousIndicatorsAccountSS->getActiveSheet();
				$IVMaliciousIndicatorsAccountS->setCellValue('A2', 'Base Feed');
				$IVMaliciousIndicatorsAccountS->setCellValue('A3', 'Base IP Feed');
				$IVMaliciousIndicatorsAccountS->setCellValue('B2', $IVPDLMaliciousDomainsPercentage ?? 0);
				$IVMaliciousIndicatorsAccountS->setCellValue('B3', $IVPDLMaliciousIPsPercentage ?? 0);
				$IVMaliciousIndicatorsAccountW = IOFactory::createWriter($IVMaliciousIndicatorsAccountSS, 'Xlsx');
				$IVMaliciousIndicatorsAccountW->save($EmbeddedIVMaliciousIndicatorsAccount);

				// Malicious Indicators Seen - Industry
				$EmbeddedIVMaliciousIndicatorsIndustry = getEmbeddedSheetFilePath('IVMaliciousIndicatorsIndustry', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$IVMaliciousIndicatorsIndustrySS = IOFactory::load($EmbeddedIVMaliciousIndicatorsIndustry);
				$IVMaliciousIndicatorsIndustryS = $IVMaliciousIndicatorsIndustrySS->getActiveSheet();
				$IVMaliciousIndicatorsIndustryS->setCellValue('A2', 'Base Feed');
				$IVMaliciousIndicatorsIndustryS->setCellValue('A3', 'Base IP Feed');
				$IVMaliciousIndicatorsIndustryS->setCellValue('B2', $IVIDLMaliciousDomainsPercentage ?? 0);
				$IVMaliciousIndicatorsIndustryS->setCellValue('B3', $IVIDLMaliciousIPsPercentage ?? 0);
				$IVMaliciousIndicatorsIndustryW = IOFactory::createWriter($IVMaliciousIndicatorsIndustrySS, 'Xlsx');
				$IVMaliciousIndicatorsIndustryW->save($EmbeddedIVMaliciousIndicatorsIndustry);

				// Risky Indicators Seen - Account
				$EmbeddedIVRiskyIndicatorsAccount = getEmbeddedSheetFilePath('IVRiskyIndicatorsAccount', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$IVRiskyIndicatorsAccountSS = IOFactory::load($EmbeddedIVRiskyIndicatorsAccount);
				$IVRiskyIndicatorsAccountS = $IVRiskyIndicatorsAccountSS->getActiveSheet();
				$IVRiskyIndicatorsAccountS->setCellValue('A2', 'Low Risk Feed');
				$IVRiskyIndicatorsAccountS->setCellValue('A3', 'Medium Risk Feed');
				$IVRiskyIndicatorsAccountS->setCellValue('A4', 'High Risk Feed');
				$IVRiskyIndicatorsAccountS->setCellValue('B2', $IVPDLUnconfirmedLowRiskPercentage ?? 0);
				$IVRiskyIndicatorsAccountS->setCellValue('B3', $IVIDLUnconfirmedMedRiskPercentage ?? 0);
				$IVRiskyIndicatorsAccountS->setCellValue('B4', $IVPDLUnconfirmedHighRiskPercentage ?? 0);
				$IVRiskyIndicatorsAccountW = IOFactory::createWriter($IVRiskyIndicatorsAccountSS, 'Xlsx');
				$IVRiskyIndicatorsAccountW->save($EmbeddedIVRiskyIndicatorsAccount);

				// Risky Indicators Seen - Industry
				$EmbeddedIVRiskyIndicatorsIndustry = getEmbeddedSheetFilePath('IVRiskyIndicatorsIndustry', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$IVRiskyIndicatorsIndustrySS = IOFactory::load($EmbeddedIVRiskyIndicatorsIndustry);
				$IVRiskyIndicatorsIndustryS = $IVRiskyIndicatorsIndustrySS->getActiveSheet();
				$IVRiskyIndicatorsIndustryS->setCellValue('A2', 'Low Risk Feed');
				$IVRiskyIndicatorsIndustryS->setCellValue('A3', 'Medium Risk Feed');
				$IVRiskyIndicatorsIndustryS->setCellValue('A4', 'High Risk Feed');
				$IVRiskyIndicatorsIndustryS->setCellValue('B2', $IVIDLUnconfirmedLowRiskPercentage ?? 0);
				$IVRiskyIndicatorsIndustryS->setCellValue('B3', $IVIDLUnconfirmedMedRiskPercentage ?? 0);
				$IVRiskyIndicatorsIndustryS->setCellValue('B4', $IVIDLUnconfirmedHighRiskPercentage ?? 0);
				$IVRiskyIndicatorsIndustryW = IOFactory::createWriter($IVRiskyIndicatorsIndustrySS, 'Xlsx');
				$IVRiskyIndicatorsIndustryW->save($EmbeddedIVRiskyIndicatorsIndustry);

				// Threat Actor Associated Traffic Seen - Account
				$EmbeddedIVThreatActorAccount = getEmbeddedSheetFilePath('IVThreatActorsAccount', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$IVThreatActorAccountSS = IOFactory::load($EmbeddedIVThreatActorAccount);
				$IVThreatActorAccountS = $IVThreatActorAccountSS->getActiveSheet();
				$IVThreatActorAccountS->setCellValue('A2', 'Associated');
				$IVThreatActorAccountS->setCellValue('A3', 'Unassociated');
				$IVThreatActorAccountS->setCellValue('B2', round($IVPDLThreatActorAssociatedTrafficSeenPercentage, 2));
				$IVThreatActorAccountS->setCellValue('B3', round(100 - $IVPDLThreatActorAssociatedTrafficSeenPercentage, 2));
				$IVThreatActorAccountW = IOFactory::createWriter($IVThreatActorAccountSS, 'Xlsx');
				$IVThreatActorAccountW->save($EmbeddedIVThreatActorAccount);

				// Threat Actor Associated Traffic Seen - Industry
				$EmbeddedIVThreatActorIndustry = getEmbeddedSheetFilePath('IVThreatActorsIndustry', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$IVThreatActorIndustrySS = IOFactory::load($EmbeddedIVThreatActorIndustry);
				$IVThreatActorIndustryS = $IVThreatActorIndustrySS->getActiveSheet();
				$IVThreatActorIndustryS->setCellValue('A2', 'Associated');
				$IVThreatActorIndustryS->setCellValue('A3', 'Unassociated');
				$IVThreatActorIndustryS->setCellValue('B2', round($IVIDLThreatActorAssociatedTrafficSeenPercentage, 2));
				$IVThreatActorIndustryS->setCellValue('B3', round(100 - $IVIDLThreatActorAssociatedTrafficSeenPercentage, 2));
				$IVThreatActorIndustryW = IOFactory::createWriter($IVThreatActorIndustrySS, 'Xlsx');
				$IVThreatActorIndustryW->save($EmbeddedIVThreatActorIndustry);

				// Threat Insight Detection - Account
				$EmbeddedIVThreatInsightAccount = getEmbeddedSheetFilePath('IVThreatInsightAccount', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$IVThreatInsightAccountSS = IOFactory::load($EmbeddedIVThreatInsightAccount);
				$IVThreatInsightAccountS = $IVThreatInsightAccountSS->getActiveSheet();
				$IVThreatInsightAccountS->setCellValue('A2', 'DNST');
				$IVThreatInsightAccountS->setCellValue('A3', 'DGA');
				$IVThreatInsightAccountS->setCellValue('B2', $IVPDLDNSTTrafficPercentage ?? 0);
				$IVThreatInsightAccountS->setCellValue('B3', $IVPDLDGATrafficPercentage ?? 0);
				$IVThreatInsightAccountW = IOFactory::createWriter($IVThreatInsightAccountSS, 'Xlsx');
				$IVThreatInsightAccountW->save($EmbeddedIVThreatInsightAccount);

				// Threat Insight Detection - Industry
				$EmbeddedIVThreatInsightIndustry = getEmbeddedSheetFilePath('IVThreatInsightIndustry', $embeddedDirectory, $embeddedFiles, $EmbeddedSheets, $SelectedTemplate['Orientation']);
				$IVThreatInsightIndustrySS = IOFactory::load($EmbeddedIVThreatInsightIndustry);
				$IVThreatInsightIndustryS = $IVThreatInsightIndustrySS->getActiveSheet();
				$IVThreatInsightIndustryS->setCellValue('A2', 'DNST');
				$IVThreatInsightIndustryS->setCellValue('A3', 'DGA');
				$IVThreatInsightIndustryS->setCellValue('B2', $IVIDLDNSTTrafficPercentage ?? 0);
				$IVThreatInsightIndustryS->setCellValue('B3', $IVIDLDGATrafficPercentage ?? 0);
				$IVThreatInsightIndustryW = IOFactory::createWriter($IVThreatInsightIndustrySS, 'Xlsx');
				$IVThreatInsightIndustryW->save($EmbeddedIVThreatInsightIndustry);

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
				// Do SOC Insight Stuff Here ....
				//
				// Skip SOC Insight Slides if Slide Number is set to 0
				if ($SelectedTemplate['SOCInsightsSlide'] != 0) {
					$Progress = $this->writeProgress($config['UUID'],$Progress,"Generating SOC Insights Slides");

					$SOCInsightSlideCount = count($SOCInsightDetails);

					// New slides to be appended after this slide number
					$SOCInsightsSlideStart = $SelectedTemplate['SOCInsightsSlide'];
					// Calculate the slide position based on above value
					$SOCInsightsSlidePosition = $SOCInsightsSlideStart-2;
					
					// Tag Numbers Start
					$SITagStart = 100;
					
					// Create Document Fragments for appending new relationships and set Starting Integers
					$xml_rels_soc_f = $xml_rels->createDocumentFragment();
					$xml_rels_soc_fstart = ($xml_rels->getElementsByTagName('Relationship')->length)+50;
					$xml_pres_soc_f = $xml_pres->createDocumentFragment();
					$xml_pres_soc_fstart = 13700;

					// Get Slide Count
					$SISlidesCount = iterator_count(new FilesystemIterator($SelectedTemplate['ExtractedDir'].'/ppt/slides'));
					// Set first slide number
					$SISlideNumber = $SISlidesCount++;
					$SIChartNumber = 50;

					foreach ($SOCInsightDetails as $SKEY => $SID) {
						// Initialise array to hold references to base excel files for SOC Insights slide
						$SOCInsightsExcelReferenceBase = [];

						if (($SOCInsightSlideCount - 1) > 0) {
							$xml_rels_soc_f->appendXML('<Relationship Id="rId'.$xml_rels_soc_fstart.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'.$SISlideNumber.'.xml"/>');
							$xml_pres_soc_f->appendXML('<p:sldId id="'.$xml_pres_soc_fstart.'" r:id="rId'.$xml_rels_soc_fstart.'"/>');
							$xml_rels_soc_fstart++;
							$xml_pres_soc_fstart++;
							copy($SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$SOCInsightsSlideStart.'.xml',$SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$SISlideNumber.'.xml');
							copy($SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$SOCInsightsSlideStart.'.xml.rels',$SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$SISlideNumber.'.xml.rels');
						} else {
							$SISlideNumber = $SOCInsightsSlideStart;
						}

						// Load Slide XML _rels
						$xml_sis = new DOMDocument('1.0', 'utf-8');
						$xml_sis->formatOutput = true;
						$xml_sis->preserveWhiteSpace = false;
						$xml_sis->load($SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$SISlideNumber.'.xml.rels');

						foreach ($xml_sis->getElementsByTagName('Relationship') as $element) {
							// Remove notes references to avoid having to create unneccessary notes resources
							if ($element->getAttribute('Type') == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide") {
								$element->remove();
							}
							if ($element->getAttribute('Type') == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart") {
								$OldChartNumber = str_replace('../charts/chart','',$element->getAttribute('Target'));
								$OldChartNumber = str_replace('.xml','',$OldChartNumber);
								copy($SelectedTemplate['ExtractedDir'].'/ppt/charts/chart'.$OldChartNumber.'.xml',$SelectedTemplate['ExtractedDir'].'/ppt/charts/chart'.$SIChartNumber.'.xml');
								copy($SelectedTemplate['ExtractedDir'].'/ppt/charts/_rels/chart'.$OldChartNumber.'.xml.rels',$SelectedTemplate['ExtractedDir'].'/ppt/charts/_rels/chart'.$SIChartNumber.'.xml.rels');
								$element->setAttribute('Target','../charts/chart'.$SIChartNumber.'.xml');

								// Load Chart XML Rels
								$xml_chart_rels = new DOMDocument('1.0', 'utf-8');
								$xml_chart_rels->formatOutput = true;
								$xml_chart_rels->preserveWhiteSpace = false;
								$xml_chart_rels->load($SelectedTemplate['ExtractedDir'].'/ppt/charts/_rels/chart'.$SIChartNumber.'.xml.rels');

								// Duplicate colours, styles & embedded excel files
								foreach ($xml_chart_rels->getElementsByTagName('Relationship') as $element_c) {
									if ($element_c->getAttribute('Type') == "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle") {
										$OldColourNumber = str_replace('colors','',$element_c->getAttribute('Target'));
										$OldColourNumber = str_replace('.xml','',$OldColourNumber);
										copy($SelectedTemplate['ExtractedDir'].'/ppt/charts/colors'.$OldColourNumber.'.xml',$SelectedTemplate['ExtractedDir'].'/ppt/charts/colors'.$SIChartNumber.'.xml');
										$element_c->setAttribute('Target','../charts/colors'.$SIChartNumber.'.xml');
									} elseif ($element_c->getAttribute('Type') == "http://schemas.microsoft.com/office/2011/relationships/chartStyle") {
										$OldStyleNumber = str_replace('style','',$element_c->getAttribute('Target'));
										$OldStyleNumber = str_replace('.xml','',$OldStyleNumber);
										copy($SelectedTemplate['ExtractedDir'].'/ppt/charts/style'.$OldStyleNumber.'.xml',$SelectedTemplate['ExtractedDir'].'/ppt/charts/style'.$SIChartNumber.'.xml');
										$element_c->setAttribute('Target','../charts/style'.$SIChartNumber.'.xml');
									} elseif ($element_c->getAttribute('Type') == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package") {
										$OldEmbeddedNumber = str_replace('../embeddings/Microsoft_Excel_Worksheet','',$element_c->getAttribute('Target'));
										$OldEmbeddedNumber = str_replace('.xlsx','',$OldEmbeddedNumber);
										copy($SelectedTemplate['ExtractedDir'].'/ppt/embeddings/Microsoft_Excel_Worksheet'.$OldEmbeddedNumber.'.xlsx',$SelectedTemplate['ExtractedDir'].'/ppt/embeddings/Microsoft_Excel_Worksheet'.$SIChartNumber.'.xlsx');
										$element_c->setAttribute('Target','../embeddings/Microsoft_Excel_Worksheet'.$SIChartNumber.'.xlsx');

										// Store the slide no. and embedded chart files for later use
										$SOCInsightsExcelReferenceBase[0][] = $SelectedTemplate['ExtractedDir'].'/ppt/embeddings/Microsoft_Excel_Worksheet'.$OldEmbeddedNumber.'.xlsx';
										$SOCInsightsExcelReference[$SKEY][] = $SelectedTemplate['ExtractedDir'].'/ppt/embeddings/Microsoft_Excel_Worksheet'.$SIChartNumber.'.xlsx';
									}
								}

								$xml_chart_rels->save($SelectedTemplate['ExtractedDir'].'/ppt/charts/_rels/chart'.$SIChartNumber.'.xml.rels');
								
								$SIChartNumber++;
							}
						}

						// $xml_sis->getElementsByTagName('Relationships')->item(0)->appendChild($xml_sis_f);
						$xml_sis->save($SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$SISlideNumber.'.xml.rels');

						// Update Tag Numbers
						$SISFile = file_get_contents($SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$SISlideNumber.'.xml');
						$SISFile = str_replace('#SITAG00', '#SITAG'.$SITagStart, $SISFile);
						file_put_contents($SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$SISlideNumber.'.xml', $SISFile);

						// Get Embedded Chart References
						if ($SKEY == 0) {
							$SOCInsightEmbeddedChart = $SOCInsightsExcelReferenceBase[0];
						} else {
							$SOCInsightEmbeddedChart = $SOCInsightsExcelReference[($SKEY)];
						}

						// Populate Charts
						foreach ([0,1] as $ChartIndex) {
							$SOCInsightEmbeddedChartRef = $SOCInsightEmbeddedChart[$ChartIndex];
							$SOCInsightEmbeddedChartSS = IOFactory::load($SOCInsightEmbeddedChartRef);
							$SOCInsightEmbeddedChartS = $SOCInsightEmbeddedChartSS->getActiveSheet();

							$SOCInsightsIndicatorsTimeSeries = $SOCInsightsCubeJSResults[$SID->insightId.'-indicatorsTimeSeries']['Body'] ?? [];
							$SOCInsightsImpactedAssetsTimeSeries = $SOCInsightsCubeJSResults[$SID->insightId.'-impactedAssetsTimeSeries']['Body'] ?? [];

							// Summerise total number of indicators by combining all distinct counts across unique timestamp.hour
							$IndicatorsTimeSeriesData = [];
							if (is_array($SOCInsightsIndicatorsTimeSeries->result->data) AND count($SOCInsightsIndicatorsTimeSeries->result->data) > 0) {
								foreach ($SOCInsightsIndicatorsTimeSeries->result->data as $ISTS) {
									$IndicatorDate = substr($ISTS->{'PortunusAggIPSummary.timestamp.hour'},0,10);
									if (isset($IndicatorsTimeSeriesData[$IndicatorDate])) {
										$IndicatorsTimeSeriesData[$IndicatorDate] += $ISTS->{'PortunusAggIPSummary.threatIndicatorDistinctCount'};
									} else {
										$IndicatorsTimeSeriesData[$IndicatorDate] = $ISTS->{'PortunusAggIPSummary.threatIndicatorDistinctCount'};
									}
								}
							}

							// Summerise total number of impacted assets by combining all distinct counts across unique timestamp.hour
							$ImpactedAssetsTimeSeriesData = [];
							if (is_array($SOCInsightsImpactedAssetsTimeSeries->result->data) AND count($SOCInsightsImpactedAssetsTimeSeries->result->data) > 0) {
								foreach ($SOCInsightsImpactedAssetsTimeSeries->result->data as $IATS) {
									$ImpactedAssetDate = substr($IATS->{'PortunusAggIPSummary.timestamp.hour'},0,10);
									if (isset($ImpactedAssetsTimeSeriesData[$ImpactedAssetDate])) {
										$ImpactedAssetsTimeSeriesData[$ImpactedAssetDate] += $IATS->{'PortunusAggIPSummary.deviceIdDistinctCount'};
									} else {
										$ImpactedAssetsTimeSeriesData[$ImpactedAssetDate] = $IATS->{'PortunusAggIPSummary.deviceIdDistinctCount'};
									}
								}
							}

							// Loop Through Indicator & Asset Time Series and add missing dates with 0 count
							$StartDate = (new DateTime($StartDimension))->format('Y-m-d');
							$EndDate = (new DateTime($EndDimension))->format('Y-m-d');
							$CurrentDate = $StartDate;

							while ($CurrentDate <= $EndDate) {
								if (!isset($IndicatorsTimeSeriesData[$CurrentDate])) {
									$IndicatorsTimeSeriesData[$CurrentDate] = 0;
								}
								if (!isset($ImpactedAssetsTimeSeriesData[$CurrentDate])) {
									$ImpactedAssetsTimeSeriesData[$CurrentDate] = 0;
								}
								// Increment Date by 1 day								
								$CurrentDate = date('Y-m-d', strtotime($CurrentDate . ' + 1 day'));
							}
							
							// Sort by date
							ksort($IndicatorsTimeSeriesData);
							ksort($ImpactedAssetsTimeSeriesData);

							switch($ChartIndex) {
								case 0:
									$RowNo = 2;
									if (is_array($ImpactedAssetsTimeSeriesData) AND count($ImpactedAssetsTimeSeriesData) > 0) {
										foreach ($ImpactedAssetsTimeSeriesData as $Date => $Count) {
											$SOCInsightEmbeddedChartS->setCellValue('A'.$RowNo, $Date);
											$SOCInsightEmbeddedChartS->setCellValue('B'.$RowNo, $Count);
											$RowNo++;
										}
									} else {
										$SOCInsightEmbeddedChartS->setCellValue('A2', date('Y-m-d'));
										$SOCInsightEmbeddedChartS->setCellValue('B2', 0);
									}
									break;
								case 1:
									$RowNo = 2;
									if (is_array($IndicatorsTimeSeriesData) AND count($IndicatorsTimeSeriesData) > 0) {
										foreach ($IndicatorsTimeSeriesData as $Date => $Count) {
											$SOCInsightEmbeddedChartS->setCellValue('A'.$RowNo, $Date);
											$SOCInsightEmbeddedChartS->setCellValue('B'.$RowNo, $Count);
											$RowNo++;
										}
									} else {
										$SOCInsightEmbeddedChartS->setCellValue('A2', date('Y-m-d'));
										$SOCInsightEmbeddedChartS->setCellValue('B2', 0);
									}
									break;
							}

							$SOCInsightEmbeddedChartW = IOFactory::createWriter($SOCInsightEmbeddedChartSS, 'Xlsx');
							$SOCInsightEmbeddedChartW->save($SOCInsightEmbeddedChartRef);
						}

						// Increment Tag Number
						$SITagStart++;
						// Decrement Slide Count
						$SOCInsightSlideCount--;
						// Increment Slide Number
						$SISlideNumber++;
					}

					// Append Elements to Core XML Files
					$xml_rels->getElementsByTagName('Relationships')->item(0)->appendChild($xml_rels_soc_f);
					// Append new slides to specific position
					$xml_pres->getElementsByTagName('sldId')->item($SOCInsightsSlidePosition)->after($xml_pres_soc_f);
				} else {
					$Progress = $this->writeProgress($config['UUID'],$Progress,"Skipping SOC Insights Slides");
				}

				//
				// Do Threat Actor Stuff Here ....
				//
				// Skip Threat Actor Slides if Slide Number is set to 0
				if ($SelectedTemplate['ThreatActorSlide'] != 0) {
					$Progress = $this->writeProgress($config['UUID'],$Progress,"Generating Threat Actor Slides");
					// New slides to be appended after this slide number
					$ThreatActorSlideStart = $SelectedTemplate['ThreatActorSlide'];
					// Calculate the slide position based on above value
					$ThreatActorSlidePosition = ($ThreatActorSlideStart-2);
		
					// Tag Numbers Start
					$TATagStart = 100;
		
					// Create Document Fragments and set Starting Integers
					$xml_rels_ta_f = $xml_rels->createDocumentFragment();
					$xml_rels_ta_fstart = ($xml_rels->getElementsByTagName('Relationship')->length)+250;
					$xml_pres_ta_f = $xml_pres->createDocumentFragment();
					$xml_pres_ta_fstart = 14700;

					// Get Slide Count
					$TASlidesCount = iterator_count(new FilesystemIterator($SelectedTemplate['ExtractedDir'].'/ppt/slides'));
					// Set first slide number
					$TASlideNumber = $TASlidesCount++;
					// Copy Blank Threat Actor Image
					copy($this->getDir()['PluginData'].'/images/logo-only.png',$SelectedTemplate['ExtractedDir'].'/ppt/media/logo-only.png');
					// Build new Threat Actor Slides & Update PPTX Resources
					$ThreatActorSlideCountIt = $ThreatActorSlideCount;
					if (isset($ThreatActorInfo)) {
						foreach  ($ThreatActorInfo as $TAI) {
							$KnownActor = $this->ThreatActors->getThreatActorConfigByName($TAI['actor_name']);
							if (($ThreatActorSlideCountIt - 1) > 0) {
								$xml_rels_ta_f->appendXML('<Relationship Id="rId'.$xml_rels_ta_fstart.'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'.$TASlideNumber.'.xml"/>');
								$xml_pres_ta_f->appendXML('<p:sldId id="'.$xml_pres_ta_fstart.'" r:id="rId'.$xml_rels_ta_fstart.'"/>');
								$xml_rels_ta_fstart++;
								$xml_pres_ta_fstart++;
								copy($SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$ThreatActorSlideStart.'.xml',$SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$TASlideNumber.'.xml');
								copy($SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$ThreatActorSlideStart.'.xml.rels',$SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$TASlideNumber.'.xml.rels');
							} else {
								$TASlideNumber = $ThreatActorSlideStart;
							}
							// Update Tag Numbers
							$TASFile = file_get_contents($SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$TASlideNumber.'.xml');
							$TASFile = str_replace('#TATAG00', '#TATAG'.$TATagStart, $TASFile);
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
							file_put_contents($SelectedTemplate['ExtractedDir'].'/ppt/slides/slide'.$TASlideNumber.'.xml', $TASFile);
							$xml_tas = new DOMDocument('1.0', 'utf-8');
							$xml_tas->formatOutput = true;
							$xml_tas->preserveWhiteSpace = false;
							$xml_tas->load($SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$TASlideNumber.'.xml.rels');
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
							$xml_tas->save($SelectedTemplate['ExtractedDir'].'/ppt/slides/_rels/slide'.$TASlideNumber.'.xml.rels');
							$TATagStart += 10;
							// Iterate slide number
							$TASlideNumber++;
							$ThreatActorSlideCountIt--;
						}
		
						// Append Elements to Core XML Files
						$xml_rels->getElementsByTagName('Relationships')->item(0)->appendChild($xml_rels_ta_f);
						// Append new slides to specific position
						$xml_pres->getElementsByTagName('sldId')->item($ThreatActorSlidePosition)->after($xml_pres_ta_f);

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

				##// Slide 37/40 - Industry Vertical Analysis - START //##

				// IVPDL - Industry Vertical Personal Dimension Label
				// IVIDL - Industry Vertical Industry Dimension Label

				// DateDiff Override
				$DateDiff = 30;

				// ** Malicious Indicators Seen ** //
				$mapping = replaceTag($mapping,'#TAG67',$IVIndustryName); // Industry Name
				$mapping = replaceTag($mapping,'#TAG68',number_abbr($DateDiff)); // Malicious Indicators - Days
				$mapping = replaceTag($mapping,'#TAG69',number_abbr($IVPDLConfirmedThreatsSeenSum)); // Malicious Indicators - Your Average - Count (Or Absolute?)
				$mapping = replaceTag($mapping,'#TAG70',$IVPDLMaliciousDomainsPercentage.'%'); // Malicious Indicators - Percentage Malicious Domains
				$mapping = replaceTag($mapping,'#TAG71',$IVPDLMaliciousIPsPercentage.'%'); // Malicious Indicators - Percentage Malicious IPs
				$mapping = replaceTag($mapping,'#TAG72',number_abbr($IVPDLConfirmedThreatsSeenSum)); // Malicious Indicators - Your Average - Count (Or Absolute?)
				$mapping = replaceTag($mapping,'#TAG73',round($IVPDLConfirmedThreatsSeenPercentage, 2).'%'); // Malicious Indicators - Your Average - Percentage (Or Absolute?)
				if ($IVPDLConfirmedThreatsSeenTrend >= 0){$MISarrow='↑';} else {$MISarrow='↓';}
				$mapping = replaceTag($mapping,'#TAG74',$MISarrow); // Arrow Up/Down
				$mapping = replaceTag($mapping,'#TAG75',number_format($IVPDLConfirmedThreatsSeenTrend, 2).'%'); // Malicious Indicators - Your Average - Percent Changed
				$mapping = replaceTag($mapping,'#TAG76',number_abbr(($IVIDLConfirmedThreatsSeenSum) / $IVIndustryPeerCount)); // Malicious Indicators - Industry Average - Count
				$mapping = replaceTag($mapping,'#TAG77',number_format($IVIDLConfirmedThreatsSeenPercentage, 2).'%'); // Malicious Indicators - Industry Average - Percentage

				// ** Risky Indicators Seen ** //
				$mapping = replaceTag($mapping,'#TAG78',number_abbr($DateDiff)); // Risky Indicators - Days
				$mapping = replaceTag($mapping,'#TAG79',number_abbr($IVPDLUnconfirmedThreatsSeenSum)); // Risky Indicators - Your Average - Count
				$mapping = replaceTag($mapping,'#TAG80',round($IVPDLUnconfirmedHighRiskPercentage, 2).'%'); // Risky Indicators - Percentage High Risk
				$mapping = replaceTag($mapping,'#TAG81',round($IVPDLUnconfirmedMedRiskPercentage, 2).'%'); // Risky Indicators - Percentage Medium Risk
				$mapping = replaceTag($mapping,'#TAG82',round($IVPDLUnconfirmedLowRiskPercentage, 2).'%'); // Risky Indicators - Percentage Low Risk
				$mapping = replaceTag($mapping,'#TAG83',number_abbr($IVPDLUnconfirmedThreatsSeenSum)); // Risky Indicators - Your Average - Count
				$mapping = replaceTag($mapping,'#TAG84',round($IVPDLUnconfirmedThreatsSeenPercentage, 2).'%'); // Risky Indicators - Your Average - Percentage
				if ($IVPDLUnconfirmedThreatsSeenTrend >= 0){$RISarrow='↑';} else {$RISarrow='↓';}
				$mapping = replaceTag($mapping,'#TAG85',$RISarrow); // Arrow Up/Down
				$mapping = replaceTag($mapping,'#TAG86',number_format($IVPDLUnconfirmedThreatsSeenTrend, 2).'%'); // Risky Indicators - Your Average - Percent Changed
				$mapping = replaceTag($mapping,'#TAG87',number_abbr(($IVIDLUnconfirmedThreatsSeenSum) / $IVIndustryPeerCount)); // Risky Indicators - Industry Average - Count
				$mapping = replaceTag($mapping,'#TAG88',number_format($IVIDLUnconfirmedThreatsSeenPercentage, 2).'%'); // Risky Indicators - Industry Average -

				// ** Threat Actor Associated Traffic Seen ** //
				$mapping = replaceTag($mapping,'#TAG89',number_abbr($DateDiff)); // Threat Actor Associated Traffic - Days
				$mapping = replaceTag($mapping,'#TAG90',number_abbr($IVThreatActorsMetrics['account'])); // Threat Actor Associated Traffic - Your Average - Count
				$mapping = replaceTag($mapping,'#TAG91',round($IVPDLThreatActorAssociatedTrafficSeenPercentage, 2).'%'); // Threat Actor Associated Traffic - Your Average
				$mapping = replaceTag($mapping,'#TAG92',number_abbr($IVThreatActorsMetrics['account'])); // Threat Actor Associated Traffic - Your Average - Count
				if ($IVPDLThreatActorAssociatedTrafficSeenTrend >= 0){$TAarrow='↑';} else {$TAarrow='↓';}
				$mapping = replaceTag($mapping,'#TAG93',$TAarrow); // Arrow Up/Down
				$mapping = replaceTag($mapping,'#TAG94',number_format($IVPDLThreatActorAssociatedTrafficSeenTrend, 2).'%'); // Threat Actor Associated Traffic - Your Average - Percent Changed
				$mapping = replaceTag($mapping,'#TAG95',round($IVPDLThreatActorAssociatedTrafficSeenPercentage, 2).'%'); // Threat Actor Associated Traffic - Your Average
				$mapping = replaceTag($mapping,'#TAG96',number_abbr($IVThreatActorsMetrics['industry'])); // Threat Actor Associated Traffic - Industry Average - Count
				$mapping = replaceTag($mapping,'#TAG97',number_format($IVIDLThreatActorAssociatedTrafficSeenPercentage, 2).'%'); // Threat Actor Associated Traffic - Industry Average

				// ** Zero Day DNS Traffic Seen ** //
				$mapping = replaceTag($mapping,'#TAG98',number_abbr($DateDiff)); // Zero Day DNS Traffic - Days
				$mapping = replaceTag($mapping,'#TAG99',number_abbr($IVPDLZeroDayDNSTrafficSeenSum)); // Zero Day DNS Traffic - Your Average - Count
				$mapping = replaceTag($mapping,'#TAG100',number_abbr($IVPDLZeroDayDNSTrafficSeenSum)); // Zero Day DNS Traffic - Your Average - Count
				if ($IVPDLZeroDayDNSTrafficSeenTrend >= 0){$ZDarrow='↑';} else {$ZDarrow='↓';}
				$mapping = replaceTag($mapping,'#TAG101',$ZDarrow); // Arrow Up/Down
				$mapping = replaceTag($mapping,'#TAG102',number_format($IVPDLZeroDayDNSTrafficSeenTrend, 2).'%'); // Zero Day DNS Traffic - Your Average - Percent Changed
				$mapping = replaceTag($mapping,'#TAG103',number_abbr(($IVIDLZeroDayDNSTrafficSeenSum) / $IVIndustryPeerCount)); // Zero Day DNS Traffic - Industry Average - Count

				// ** Threat Insight Detection ** //
				$mapping = replaceTag($mapping,'#TAG104',number_abbr($DateDiff)); // Threat Insight Detection - Days
				$mapping = replaceTag($mapping,'#TAG105',number_abbr($IVPDLThreatInsightDetectionSum)); // Threat Insight Detection - Your Average - Count
				$mapping = replaceTag($mapping,'#TAG106',round($IVPDLDNSTTrafficPercentage, 2).'%'); // Threat Insight Detection - Your Average - DNST Count
				$mapping = replaceTag($mapping,'#TAG107',round($IVPDLDGATrafficPercentage, 2).'%'); // Threat Insight Detection - Your Average - DGA Count
				$mapping = replaceTag($mapping,'#TAG108',number_abbr($IVPDLThreatInsightDetectionSum)); // Threat Insight Detection - Your Average - Count
				$mapping = replaceTag($mapping,'#TAG109',round($IVPDLThreatInsightDetectionPercentage, 2).'%'); // Threat Insight Detection - Your Average - Percentage
				if ($IVPDLThreatInsightDetectionTrend >= 0){$TIDarrow='↑';} else {$TIDarrow='↓';}
				$mapping = replaceTag($mapping,'#TAG110',$TIDarrow); // Arrow Up/Down
				$mapping = replaceTag($mapping,'#TAG111',number_format($IVPDLThreatInsightDetectionTrend, 2).'%'); // Threat Insight Detection - Your Average - Percent Changed
				$mapping = replaceTag($mapping,'#TAG112',number_abbr(($IVIDLThreatInsightDetectionSum) / $IVIndustryPeerCount)); // Threat Insight Detection - Industry Average - Count
				$mapping = replaceTag($mapping,'#TAG113',number_format($IVIDLThreatInsightDetectionPercentage, 2).'%'); // Threat Insight Detection - Industry Average - Percentage

				##// Slide 37/40 - Industry Vertical Analysis - END //##
				
				##// Slide 51 - Token Calculator //##
				$mapping = replaceTag($mapping,'#TAG114',number_abbr($SecurityTokenCategories['Total']*1.2)); // Total Security Tokens + 20%
				$mapping = replaceTag($mapping,'#TAG115',number_abbr($SecurityTokenCategories['Cloud Asset Protection']['Total']*1.2)); // Threat Defense Cloud Asset Protection Tokens + 20%
				$mapping = replaceTag($mapping,'#TAG116',number_abbr($SecurityTokenCategories['Threat Defense for NIOS']['Total']*1.2)); // Threat Defense for NIOS Tokens + 20%
				$mapping = replaceTag($mapping,'#TAG117',number_abbr($SecurityTokenCategories['SOC Insights']['Total']*1.2)); // Threat Defense SOC Insights Tokens + 20%
				$mapping = replaceTag($mapping,'#TAG118',number_abbr($SecurityTokenCategories['Lookalikes']['Total']*1.2)); // Threat Defense Lookalike Tokens + 20%
				$mapping = replaceTag($mapping,'#TAG119',number_abbr($SecurityTokenCategories['Dossier']['Total']*1.2)); // Threat Defense Dossier Tokens + 20%
				$mapping = replaceTag($mapping,'#TAG120',number_abbr($ReportingTokenCategories['Total']*1.2)); // Total Reporting Tokens + 20%
				$mapping = replaceTag($mapping,'#TAG121',number_abbr($ReportingTokenCategories['Ecosystem']*1.2)); // Total Ecosystem Tokens + 20%
				$mapping = replaceTag($mapping,'#TAG122',number_abbr($ReportingTokenCategories['S3']*1.2)); // Total S3 Tokens + 20%

				// Calculate Estimated Reporting Tokens, where it may not be enabled today. This is a based on DNS Activity + Security Activity, 10M / 40 tokens
				$TotalReportingEventsCount = $DNSActivityCount + $SecurityEventsCount;
				$EstimatedReportingTokens = ceil(($TotalReportingEventsCount / 10000000) * 40);
				$mapping = replaceTag($mapping,'#TAG123',number_abbr($EstimatedReportingTokens*1.2)); // Estimated Reporting Tokens + 20%

				##// Slide 16+ - SOC Insight Details
				$SITagStart = 100;
				if (isset($SOCInsightDetails)) {
					foreach ($SOCInsightDetails as $SID) {
						$SOCInsightCubeResponseGeneral = $SOCInsightsCubeJSResults[$SID->insightId.'-general']['Body'] ?? [];
						$SOCInsightCubeResponseAssets = $SOCInsightsCubeJSResults[$SID->insightId.'-assetCounts']['Body'] ?? [];
						$SOCInsightIndicatorsBlockedCounts = $SOCInsightsIndicatorsBlockedCounts[$SID->insightId] ?? [];
						$SOCInsightIndicatorsDetails = $SOCInsightsIndicatorsDetails[$SID->insightId]->result ?? [];
						$SOCInsightAssetAccessingCount = ($SOCInsightsCubeJSResults[$SID->insightId.'-assetsAccessingQIP']['Body']->result->data[0]->{'PortunusAggIPSummary.qipDistinctCount'} ?? 0) + ($SOCInsightsCubeJSResults[$SID->insightId.'-assetsAccessingDeviceID']['Body']->result->data[0]->{'PortunusAggIPSummary.deviceIdDistinctCount'} ?? 0);

						$MassSpreading = '';
						$PersistentThreat = '';
						$TotalEvents = 0;
						foreach ($SOCInsightCubeResponseGeneral->result->data as $InsightGeneral) {
							if ($InsightGeneral->{'InsightDetails.spreading'}) {
								$MassSpreading = 'Mass Spreading';
							}
							if ($InsightGeneral->{'InsightDetails.persistent'}) {
								$PersistentThreat = 'Persistent Threat';
							}
							$TotalEvents += $InsightGeneral->{'InsightDetails.numEvents'} ?? 0;
						}

						$startedAt = new DateTime($SID->startedAt);
						$mostRecentAt = new DateTime($SID->mostRecentAt);
						$ActivePeriod = $startedAt->diff($mostRecentAt);
						$ActivePeriodDays = $ActivePeriod->days ?? 0;

						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'01',$SID->threatType ?? ''); // SOC Insight Name
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'02',$ActivePeriodDays); // Active Period
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'03',$SID->startedAt ?? ''); // Insight Creation Date
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'04',$SID->mostRecentAt ?? ''); // Last Observed Date
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'05',$SID->threatType ?? ''); // Insight Type
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'06',$SID->tClass ?? ''); // Insight Class
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'07',$SID->tFamily ?? ''); // Insight Family
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'08',$MassSpreading); // Mass Spreading Y/N
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'09',$PersistentThreat); // Persistent Threat Y/N
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'10',$SOCInsightAssetAccessingCount); // Qty Asset Accessing
						// $mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'11',$SID->eventsNotBlockedCount ?? 0); // Events Not Blocked
						// $mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'12',$SID->eventsBlockedCount ?? 0); // Indicators Blocked
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'11',$SOCInsightIndicatorsBlockedCounts->notBlocked ?? 0); // Indicators Not Blocked
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'12',$SOCInsightIndicatorsBlockedCounts->blocked ?? 0); // Indicators Blocked
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'13',$SOCInsightCubeResponseAssets->result->data[0]->{'PortunusAggIPSummary_ch.totalAssetCount'} ?? 0); // Total Assets
						$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.'14',$SOCInsightIndicatorsBlockedCounts->total ?? 0); // Total Indicators

						// ** Indicator Table ** //
						$SOCInsightIndicatorDetailStartTag = 15;
						foreach ([0,1] as $IndicatorItem) {
							if (isset($SOCInsightIndicatorsDetails[$IndicatorItem])) {
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.$SOCInsightIndicatorDetailStartTag,$SOCInsightIndicatorsDetails[$IndicatorItem]->indicator ?? ''); // Indicator
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+1),$SOCInsightIndicatorsDetails[$IndicatorItem]->action ?? ''); // Blocked/Not Blocked Label
								if (isset($SOCInsightIndicatorsDetails[$IndicatorItem]->action)) {
									if ($SOCInsightIndicatorsDetails[$IndicatorItem]->action == 'Blocked') {
										$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+2),'Indicator has been added To a Threat Feed and blocked.'); // Blocked Label
										$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+3),''); // Priority Label
									} else {
										$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+2),'Indicator is not blocked. Recommend to block.'); // Not Blocked Label
										$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+3),'High Priority'); // Priority Label
									}
								} else {
									$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+2),'N/A'); // Not Blocked Label
									$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+3),'N/A'); // Priority Label
								}

								// Map Threat & Confidence Level to Low, Medium, High
								$ThreatLevel = $CriticalalityMapping[$SOCInsightIndicatorsDetails[$IndicatorItem]->threatLevelMax ?? 5] ?? '';
								$ConfidenceLevel = $CriticalalityMapping[$SOCInsightIndicatorsDetails[$IndicatorItem]->confidenceLevelMax ?? 5] ?? '';

								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+4),$SOCInsightIndicatorsDetails[$IndicatorItem]->qipDistinctCount); // Asset Qty
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+5),$ThreatLevel); // Threat Level
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+6),$ConfidenceLevel); // Confidence
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+7),$SOCInsightIndicatorsDetails[$IndicatorItem]->timestampMax ?? ''); // Last Observation
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+8),$SOCInsightIndicatorsDetails[$IndicatorItem]->timestampMin ?? ''); // First Observation
							} else {
								// No Indicator Data, clear out the tags
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.$SOCInsightIndicatorDetailStartTag,''); // Indicator
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+1),''); // Blocked/Not Blocked Label
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+2),''); // Blocked Label
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+3),''); // Priority Label
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+4),''); // Asset Qty
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+5),''); // Threat Level
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+6),''); // Confidence
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+7),''); // Last Observation
								$mapping = replaceTag($mapping,'#SITAG'.$SITagStart.($SOCInsightIndicatorDetailStartTag+8),''); // First Observation

							}
							$SOCInsightIndicatorDetailStartTag += 9;
						}
						
						$SITagStart++;
					}
				}

				##// Slide 32/34 - Threat Actors
				// This is where the Threat Actor Tag replacement occurs
				// Set Tag Start Number
				$TATagStart = 100;
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
		
						$mapping = replaceTag($mapping,'#TATAG'.$TATagStart.'01',ucwords($TAI['actor_name'])); // Threat Actor Name
						$mapping = replaceTag($mapping,'#TATAG'.$TATagStart.'02',$TAI['actor_description']); // Threat Actor Description
						$mapping = replaceTag($mapping,'#TATAG'.$TATagStart.'03',$IndicatorCount); // Number of Observed IOCs
						$mapping = replaceTag($mapping,'#TATAG'.$TATagStart.'04',$IndicatorsNotObserved); // Number of Observed IOCs not observed
						$mapping = replaceTag($mapping,'#TATAG'.$TATagStart.'05',$TAI['related_count']); // Number of Related IOCs
						$mapping = replaceTag($mapping,'#TATAG'.$TATagStart.'06',$DaysAhead); // Discovered X Days Ahead
						$mapping = replaceTag($mapping,'#TATAG'.$TATagStart.'07',$FirstSeen->format('d/m/y')); // First Detection Date
						$mapping = replaceTag($mapping,'#TATAG'.$TATagStart.'08',$LastSeen->format('d/m/y')); // Last Detection Date
						$mapping = replaceTag($mapping,'#TATAG'.$TATagStart.'09',$ProtectedFor); // Protected X Days
						$mapping = replaceTag($mapping,'#TATAG'.$TATagStart.'10',$ExampleDomain); // Example Domain
						$TATagStart += 10;
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
			'Total' => ($Total * 19) + 35,
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