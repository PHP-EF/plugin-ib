<?php
class ibPortal extends ibPlugin {
	protected $APIKey;
	protected $Realm;

	public function __construct() {
		parent::__construct();
	}

	public function SetCSPConfiguration($APIKey = null,$Realm = "US") {
		$this->Realm = $Realm;
		if (isset($_COOKIE['crypt'])) {
			$this->APIKey = decrypt($_COOKIE['crypt'],$this->config->get("Security","salt"));
		} else {
			$this->APIKey = $APIKey;
		}
		if ($this->CheckAPIKey()) {
			return true;
		}
	}

	public function GetCSPConfiguration($Uri = null) {
		$CSPHeaders = array(
			'Authorization' => "Token {$this->APIKey}",
			'Content-Type' => "application/json"
		);
	  
		$ErrorOnEmpty = true;
	  
		if ($Uri == null || strpos($Uri,"https://csp.") === FALSE) {
		  if ($this->Realm == "US") {
			$Url = "https://csp.infoblox.com/".$Uri;
		  } elseif ($this->Realm == "EU") {
			$Url = "https://csp.eu.infoblox.com/".$Uri;
		  } else {
			echo 'Error. Invalid Realm';
			return false;
		  }
		} else {
		  $Url = $Uri;
		}
	  
		$Options = array(
		  'timeout' => $this->config->get("System","CURL-Timeout"),
		  'connect_timeout' => $this->config->get("System","CURL-ConnectTimeout")
		);
	  
		return array(
		  "APIKey" => $this->APIKey,
		  "Realm" => $this->Realm,
		  "Url" => $Url,
		  "Options" => $Options,
		  "Headers" => $CSPHeaders
		);
	}
	  
	function QueryCSPMultiRequestBuilder($Method = 'GET', $Uri = '', $Data = null, $Id = "") {
		return array(
			"Id" => $Id,
			"Method" => $Method,
			"Uri" => $Uri,
			"Data" => $Data
		);
	}
	
	function QueryCSPMulti($MultiQuery) {
		$CSPConfig = $this->GetCSPConfiguration(null);
		
		// Prepare the requests
		$requests = [];
		foreach ($MultiQuery as $Multi) {
			if (isset($Multi['Data'])) { $Data = $Multi['Data']; } else { $Data = null; }
			$requests[] = array(
				"id" => $Multi['Id'],
				"type" => constant('Requests::' . strtoupper($Multi['Method'])),
				"url" => $CSPConfig['Url'].$Multi['Uri'],
				"data" => $Data,
				"headers" => $CSPConfig['Headers'],
				"options" => $CSPConfig['Options']
			);
		}
		// Send the requests simultaneously
		$responses = Requests::request_multiple($requests);
		
		$Results = [];
		$IdStepIn = 0;
		foreach ($responses as $index => $response) {
			if ($requests[$index]['id'] != "") {
				$Id = $requests[$index]['id'];
			} else {
				$Id = $IdStepIn;
				$IdStepIn++;
			}
			$Results[$Id] = array(
				'Response' => $response,
				'Body' => json_decode($response->body)
			);
		}
		return $Results;
	}
	
	function QueryCSP($Method, $Uri, $Data = null) {
		$CSPConfig = $this->GetCSPConfiguration($Uri);
		try {
			switch ($Method) {
				case 'get':
					$Result = Requests::get($CSPConfig['Url'], $CSPConfig['Headers'], $CSPConfig['Options']);
					break;
				case 'post':
					if ($Data != null) {
						$Result = Requests::post($CSPConfig['Url'], $CSPConfig['Headers'], json_encode($Data,JSON_UNESCAPED_SLASHES), $CSPConfig['Options']);
					} else {
						$Result = Requests::post($CSPConfig['Url'], $CSPConfig['Headers'], $Data, $CSPConfig['Options']);
					}
					break;
				case 'put':
					$Result = Requests::put($CSPConfig['Url'], $CSPConfig['Headers'], json_encode($Data,JSON_UNESCAPED_SLASHES), $CSPConfig['Options']);
					break;
				case 'patch':
					$Result = Requests::patch($CSPConfig['Url'], $CSPConfig['Headers'], json_encode($Data,JSON_UNESCAPED_SLASHES), $CSPConfig['Options']);
					break;
				case 'delete':
					$Result = Requests::delete($CSPConfig['Url'], $CSPConfig['Headers'], $CSPConfig['Options']);
					break;
			}
		} catch (Exception $e) {
			return array(
				'Status' => 'Error',
				'Error' => $e->getMessage()
			);
		}
		
		$LogArr = array(
			"Method" => $Method,
			"Url" => $CSPConfig['Url'],
			"Options" => $CSPConfig['Options']
		);
		
		if ($Result) {
			switch ($Result->status_code) {
				case '401':
					$LogArr['Error'] = "Invalid API Key.";
					$this->logging->writeLog("CSP","Failed to authenticate to the CSP","debug",$LogArr);
					$this->api->setAPIResponse('Error','Invalid API Key');
					return false;
				default:
					$Output = json_decode($Result->body);
					$this->logging->writeLog("CSP","Queried the CSP","debug",$LogArr);
					return $Output;
				}
		} elseif ($ErrorOnEmpty) {
			echo "Warning. No results from API.".$CSPConfig->Url;
		}
	}
	
	public function QueryCubeJSMulti($MultiQuery) {
		$BuildQuery = [];
		foreach ($MultiQuery as $Id => $Query) {
			$BuildQuery[] = $this->QueryCSPMultiRequestBuilder("get","api/cubejs/v1/query?query=".urlencode($Query),null,$Id);
		}
		$Results = $this->QueryCSPMulti($BuildQuery);
		return $Results;
	}
	
	public function QueryCubeJS($Query) {
		$BuildQuery = urlencode($Query);
		$Result = $this->QueryCSP("get","api/cubejs/v1/query?query=".$BuildQuery);
		return $Result;
	}

	public function CheckAPIKey() {
		if ($this->APIKey) {
			if ($this->GetCSPCurrentUser()) {
				return true;
			}
		}
		return false;
	}
	
	public function GetCSPCurrentUser() {
		$UserInfo = $this->QueryCSP("get","v2/current_user");
		return $UserInfo;
	}
}