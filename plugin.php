<?php
// **
// USED TO DEFINE PLUGIN INFORMATION & CLASS
// **

// PLUGIN INFORMATION - This should match what is in plugin.json
$GLOBALS['plugins']['IB-Tools'] = [ // Plugin Name
	'name' => 'IB-Tools', // Plugin Name
	'author' => 'TehMuffinMoo', // Who wrote the plugin
	'category' => 'Infoblox Tools', // One to Two Word Description
	'link' => 'https://github.com/php-ef/plugin-ib', // Link to plugin info
	'version' => '1.3.3', // SemVer of plugin
	'image' => 'logo.png', // 1:1 non transparent image for plugin
	'settings' => true, // does plugin need a settings modal?
	'api' => '/api/plugin/ib/settings', // api route for settings page
];

// Include IB-Tools Classes
foreach (glob(__DIR__.'/classes/*.php') as $function) {
    require_once $function; // Include each PHP file
}

class ibPlugin extends phpef {
	public $SecurityAssessment;
	public $sql;

	public function __construct() {
	   parent::__construct();
	   $dbFile = dirname(__DIR__,2). DIRECTORY_SEPARATOR . 'config' . DIRECTORY_SEPARATOR . 'IB-Tools.db';
	   $this->sql = new PDO("sqlite:$dbFile");
	}

	public function _pluginGetSettings() {
		return array(
			'Plugin Settings' => array(
				$this->settingsOption('auth', 'ACL-SECURITYASSESSMENT', ['label' => 'Security Assessment ACL']),
				$this->settingsOption('auth', 'ACL-THREATACTORS', ['label' => 'Threat Actors ACL']),
				$this->settingsOption('auth', 'ACL-LICENSEUSAGE', ['label' => 'License Utilization ACL']),
				$this->settingsOption('auth', 'ACL-CLOUDASSESSMENT', ['label' => 'Cloud Assessment ACL']),
				$this->settingsOption('auth', 'ACL-CONFIG', ['label' => 'Configuration Admin ACL']),
				$this->settingsOption('auth', 'ACL-REPORTING', ['label' => 'Reporting ACL']),
			)
		);
	}

	public function getDir() {
		return array(
			'Files' => dirname(__DIR__, 3) . DIRECTORY_SEPARATOR . 'files',
			'Assets' => dirname(__DIR__, 3) . DIRECTORY_SEPARATOR . 'assets',
			'PluginData' => __DIR__ . DIRECTORY_SEPARATOR . 'data'
		);
	}

	function timeAgo($timestamp) {
		$time = strtotime($timestamp);
		$currentTime = time();
		$diff = $currentTime - $time;
		
		$seconds = $diff;
		$minutes = round($diff / 60);
		$hours = round($diff / 3600);
		$remainingMinutes = floor(($diff % 3600) / 60);
		$days = round($diff / 86400);
		$remainingHours = floor(($diff % 86400) / 3600);
		
		if ($seconds < 60) {
			return $seconds . 's';
		} elseif ($minutes < 60) {
			return $minutes . 'm';
		} elseif ($hours < 24) {
			return $hours . 'h ' . $remainingMinutes . 'm';
		} else {
			return $days . 'd ' . $remainingHours . 'h';
		}
	}
		
}