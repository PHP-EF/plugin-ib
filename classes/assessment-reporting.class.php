<?php
class AssessmentReporting extends ibPlugin {
    public function __construct() {
		parent::__construct();
        // Create or open the SQLite database
        $this->createReportingTables();
		$this->createSecurityMetricsTables();
    }

	private function createSecurityMetricsTables() {
	    // Create security assessments anonymised metrics table if it doesn't exist
		$this->sql->exec("CREATE TABLE IF NOT EXISTS anonymised_metrics_security (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			date_start DATETIME,
			date_end DATETIME,
			dns_requests INTEGER,
			security_events_high_risk INTEGER,
			security_events_medium_risk INTEGER,
			security_events_low_risk INTEGER,
			security_events_doh INTEGER,
			security_events_zero_day INTEGER,
			security_events_suspicious INTEGER,
			security_events_newly_observed_domains INTEGER,
			security_events_dga INTEGER,
			security_events_tunnelling INTEGER,
			security_insights INTEGER,
			security_threat_actors INTEGER,
			web_unique_applications INTEGER,
			web_high_risk_categories INTEGER,
			lookalikes_custom_domains INTEGER
		)");
	}

	private function createReportingTables() {
	    // Create assessments table if it doesn't exist
		$this->sql->exec("CREATE TABLE IF NOT EXISTS reporting_assessments (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			type TEXT,
			userid INTEGER,
			apiuser TEXT,
			customer TEXT,
			realm TEXT,
			created DATETIME,
			uuid TEXT,
			status TEXT
		)");
	}

	// Assessment Reports
	public function getReportById($id) {
		$stmt = $this->sql->prepare("SELECT * FROM reporting_assessments WHERE id = :id");
		$stmt->execute([':id' => $id]);
		$report = $stmt->fetch(PDO::FETCH_ASSOC);
		if ($report) {
			return $report;
		} else {
			return false;
		}
	}
	
	public function getReportByUuid($uuid) {
		$stmt = $this->sql->prepare("SELECT * FROM reporting_assessments WHERE uuid = :uuid");
		$stmt->execute([':uuid' => $uuid]);
		$report = $stmt->fetch(PDO::FETCH_ASSOC);
		if ($report) {
		 	 return $report;
		} else {
		 	 return false;
		}
	}
	
	public function newReportEntry($type,$apiuser,$customer,$realm,$uuid,$status) {
		$stmt = $this->sql->prepare("INSERT INTO reporting_assessments (type, apiuser, customer, realm, created, uuid, status) VALUES (:type, :apiuser, :customer, :realm, :created, :uuid, :status)");
		$stmt->execute([':type' => $type,':apiuser' => $apiuser,':customer' => $customer,':realm' => $realm,':created' => date('Y-m-d H:i:s'),':uuid' => $uuid,':status' => $status]);
		return $this->sql->lastInsertId();
	}
	
	public function updateReportEntry($id,$type,$apiuser,$customer,$realm,$uuid,$status) {
		if ($this->getReportById($id)) {
			$prepare = [];
			$execute = [];
			$execute[':id'] = $id;
			if ($type !== null) {
				$prepare[] = 'type = :type';
				$execute[':type'] = $type;
			}
			if ($apiuser !== null) {
				$prepare[] = 'apiuser = :apiuser';
				$execute[':apiuser'] = $apiuser;
			}
			if ($customer !== null) {
				$prepare[] = 'customer = :customer';
				$execute[':customer'] = $customer;
			}
			if ($realm !== null) {
				$prepare[] = 'realm = :realm';
				$execute[':realm'] = $realm;
			}
			if ($uuid !== null) {
				$prepare[] = 'uuid = :uuid';
				$execute[':uuid'] = $uuid;
			}
			if ($status !== null) {
				$prepare[] = 'status = :status';
				$execute[':status'] = $status;
			}
			$stmt = $this->sql->prepare('UPDATE reporting_assessments SET '.implode(", ",$prepare).' WHERE id = :id');
			$stmt->execute($execute);
			return array(
				'Status' => 'Success',
				'Message' => 'Report Record updated successfully'
			);
		} else {
			return array(
				'Status' => 'Error',
				'Message' => 'Report Record does not exist'
			);
		}
	}
	
	public function updateReportEntryStatus($uuid,$status) {
		$stmt = $this->sql->prepare('UPDATE reporting_assessments SET status = :status WHERE uuid = :uuid');
		$stmt->execute([':uuid' => $uuid,':status' => $status]);
	}
	
	public function getAssessmentReports($granularity,$filters,$start = null,$end = null) {
		$execute = [];
		$Select = $this->reporting->sqlSelectByGranularity($granularity,'created','reporting_assessments',$start,$end);
		if ($granularity == 'custom') {
			if ($start != null && $end != null) {
				$StartDateTime = (new DateTime($start))->format('Y-m-d H:i:s');
				$EndDateTime = (new DateTime($end))->format('Y-m-d H:i:s');
				$execute[':start'] = $StartDateTime;
				$execute[':end'] = $EndDateTime;
			}
		}
		if ($filters['type'] != 'all') {
			$Select = $Select.' AND type = :type';
			$execute[':type'] = $filters['type'];
		}
		if ($filters['realm'] != 'all') {
			$Select = $Select.' AND realm = :realm';
			$execute[':realm'] = $filters['realm'];
		}
		if ($filters['user'] != 'all') {
			$Select = $Select.' AND apiuser = :apiuser';
			$execute[':apiuser'] = $filters['user'];
		}
		if ($filters['customer'] != 'all') {
			$Select = $Select.' AND customer = :customer';
			$execute[':customer'] = $filters['customer'];
		}
		if (isset($Select)) {
			$Select .= " ORDER BY created DESC";
		  try {
			$stmt = $this->sql->prepare($Select);
			$stmt->execute($execute);
			return $stmt->fetchAll(PDO::FETCH_ASSOC);
		  } catch (PDOException $e) {
				return array(
					'Status' => 'Error',
					'Message' => $e
				);
		  }
		} else {
			return array(
				'Status' => 'Error',
				'Message' => 'Invalid Granularity'
			);
		}
	}
	
	public function getAssessmentReportsSummary() {
		// Include grouping by type
		//$stmt = $this->sql->prepare('SELECT type, SUM(CASE WHEN DATE(created) = DATE("now") THEN 1 ELSE 0 END) AS count_today, SUM(CASE WHEN strftime("%Y-%m", created) = strftime("%Y-%m", "now") THEN 1 ELSE 0 END) AS count_this_month, SUM(CASE WHEN strftime("%Y", created) = strftime("%Y", "now") THEN 1 ELSE 0 END) AS count_this_year, COUNT(DISTINCT CASE WHEN DATE(created) = DATE("now") THEN apiuser ELSE NULL END) AS unique_apiusers_today, COUNT(DISTINCT CASE WHEN strftime("%Y-%m", created) = strftime("%Y-%m", "now") THEN apiuser ELSE NULL END) AS unique_apiusers_this_month, COUNT(DISTINCT CASE WHEN strftime("%Y", created) = strftime("%Y", "now") THEN apiuser ELSE NULL END) AS unique_apiusers_this_year, COUNT(DISTINCT CASE WHEN DATE(created) = DATE("now") THEN customer ELSE NULL END) AS unique_customers_today, COUNT(DISTINCT CASE WHEN strftime("%Y-%m", created) = strftime("%Y-%m", "now") THEN customer ELSE NULL END) AS unique_customers_this_month, COUNT(DISTINCT CASE WHEN strftime("%Y", created) = strftime("%Y", "now") THEN customer ELSE NULL END) AS unique_customers_this_year FROM reporting_assessments GROUP BY type UNION ALL SELECT "Total" AS type, SUM(CASE WHEN DATE(created) = DATE("now") THEN 1 ELSE 0 END) AS count_today, SUM(CASE WHEN strftime("%Y-%m", created) = strftime("%Y-%m", "now") THEN 1 ELSE 0 END) AS count_this_month, SUM(CASE WHEN strftime("%Y", created) = strftime("%Y", "now") THEN 1 ELSE 0 END) AS count_this_year, COUNT(DISTINCT CASE WHEN DATE(created) = DATE("now") THEN apiuser ELSE NULL END) AS unique_apiusers_today, COUNT(DISTINCT CASE WHEN strftime("%Y-%m", created) = strftime("%Y-%m", "now") THEN apiuser ELSE NULL END) AS unique_apiusers_this_month, COUNT(DISTINCT CASE WHEN strftime("%Y", created) = strftime("%Y", "now") THEN apiuser ELSE NULL END) AS unique_apiusers_this_year, COUNT(DISTINCT CASE WHEN DATE(created) = DATE("now") THEN customer ELSE NULL END) AS unique_customers_today, COUNT(DISTINCT CASE WHEN strftime("%Y-%m", created) = strftime("%Y-%m", "now") THEN customer ELSE NULL END) AS unique_customers_this_month, COUNT(DISTINCT CASE WHEN strftime("%Y", created) = strftime("%Y", "now") THEN customer ELSE NULL END) AS unique_customers_this_year FROM reporting_assessments;');
		$stmt = $this->sql->prepare('SELECT "Total" AS type, SUM(CASE WHEN DATE(created) = DATE("now") THEN 1 ELSE 0 END) AS count_today, SUM(CASE WHEN strftime("%Y-%m", created) = strftime("%Y-%m", "now") THEN 1 ELSE 0 END) AS count_this_month, SUM(CASE WHEN strftime("%Y", created) = strftime("%Y", "now") THEN 1 ELSE 0 END) AS count_this_year, COUNT(DISTINCT CASE WHEN DATE(created) = DATE("now") THEN apiuser ELSE NULL END) AS unique_apiusers_today, COUNT(DISTINCT CASE WHEN strftime("%Y-%m", created) = strftime("%Y-%m", "now") THEN apiuser ELSE NULL END) AS unique_apiusers_this_month, COUNT(DISTINCT CASE WHEN strftime("%Y", created) = strftime("%Y", "now") THEN apiuser ELSE NULL END) AS unique_apiusers_this_year, COUNT(DISTINCT CASE WHEN DATE(created) = DATE("now") THEN customer ELSE NULL END) AS unique_customers_today, COUNT(DISTINCT CASE WHEN strftime("%Y-%m", created) = strftime("%Y-%m", "now") THEN customer ELSE NULL END) AS unique_customers_this_month, COUNT(DISTINCT CASE WHEN strftime("%Y", created) = strftime("%Y", "now") THEN customer ELSE NULL END) AS unique_customers_this_year FROM reporting_assessments;');
		$stmt->execute();
		return $stmt->fetchAll(PDO::FETCH_ASSOC);
	}
	
	public function getAssessmentReportsStats($granularity,$filters,$start,$end) {
		$data = $this->getAssessmentReports($granularity,$filters,$start,$end);
		$summary = $this->summarizeByTypeAndDate($data, $granularity);
		return $summary;
	}
	
	// Function to summarize data by type and date
	private function summarizeByTypeAndDate($data, $granularity) {
		$summary = [];
		foreach ($data as $item) {
			$type = $item['type'];
			$dateKey = $this->reporting->summerizeDateByGranularity($item,$granularity,'created');
			if (!isset($summary[$type])) {
				$summary[$type] = [];
			}
			if (!isset($summary[$type][$dateKey])) {
				$summary[$type][$dateKey] = 0;
			}
			$summary[$type][$dateKey]++;
		}
		return $summary;
	}

	// Anonymised Security Metrics
	public function newSecurityMetricsEntry($data) {
		// Prepare the SQL statement to insert a new entry based on the provided data, which may not include all fields
		$fields = [];
		$placeholders = [];
		$values = [];

		$StartDate = $data['date_start'];
		$EndDate = $data['date_end'];
		$DateDiff = $EndDate->getTimestamp() - $StartDate->getTimestamp();
		$data['date_start'] = date_format($data['date_start'], 'Y-m-d H:i:s');
		$data['date_end'] = date_format($data['date_end'], 'Y-m-d H:i:s');

		foreach ($data as $key => $value) {
			$fields[] = $key;
			$placeholders[] = ':' . $key;

			if ($key != 'date_start' && $key != 'date_end') {
				// Normalise the value to an hourly value based on the difference between start and end dates $StartDate and $EndDate
				$hours = max(1, round($DateDiff / 3600)); // Ensure at least 1 hour
				$outvalue = round($value / $hours, 2); // Normalise to hourly value
			} else {
				$outvalue = $value; // Keep date values as they are
			}

			$values[':' . $key] = $outvalue;
		}
		$sql = "INSERT INTO anonymised_metrics_security (" . implode(',', $fields) . ") VALUES (" . implode(',', $placeholders) . ")";
		$stmt = $this->sql->prepare($sql);
		$stmt->execute($values);
		return $this->sql->lastInsertId();
	}

	public function getAnonymisedMetricsSecurity($granularity,$filters,$start = null,$end = null) {
		$execute = [];
		$Select = $this->reporting->sqlSelectByGranularity($granularity,'date_end','anonymised_metrics_security',$start,$end);
		if ($granularity == 'custom') {
			if ($start != null && $end != null) {
				$StartDateTime = (new DateTime($start))->format('Y-m-d H:i:s');
				$EndDateTime = (new DateTime($end))->format('Y-m-d H:i:s');
				$execute[':start'] = $StartDateTime;
				$execute[':end'] = $EndDateTime;
			}
		}
		if (isset($Select)) {
			try {
				$stmt = $this->sql->prepare($Select);
				$stmt->execute($execute);
				return $stmt->fetchAll(PDO::FETCH_ASSOC);
			} catch (PDOException $e) {
				return array(
					'Status' => 'Error',
					'Message' => $e
				);
			}
		} else {
			return array(
				'Status' => 'Error',
				'Message' => 'Invalid Granularity'
			);
		}
	}


	function getAnonymisedMetricsSecurityAverages($granularity, $startDateTime, $endDateTime) {
	    $results = $this->getAnonymisedMetricsSecurity($granularity, [], $startDateTime, $endDateTime);
	    $sums = [];
	    $recordCount = 0;

		if ($granularity != 'custom') {
			switch ($granularity) {
				case 'today':
					$startDateTime = date('Y-m-d 00:00:00');
					$endDateTime = date('Y-m-d 23:59:59');
					break;
				case 'thisWeek':
					$startDateTime = date('Y-m-d 00:00:00', strtotime('monday this week'));
					$endDateTime = date('Y-m-d 23:59:59', strtotime('sunday this week'));
					break;
				case 'thisMonth':
					$startDateTime = date('Y-m-01 00:00:00');
					$endDateTime = date('Y-m-t 23:59:59');
					break;
				case 'thisYear':
					$startDateTime = date('Y-01-01 00:00:00');
					$endDateTime = date('Y-12-31 23:59:59');
					break;
				case 'last30Days':
					$startDateTime = date('Y-m-d 00:00:00', strtotime('-30 days'));
					$endDateTime = date('Y-m-d 23:59:59');
					break;
				case 'lastMonth':
					$startDateTime = date('Y-m-01 00:00:00', strtotime('first day of last month'));
					$endDateTime = date('Y-m-t 23:59:59', strtotime('last day of last month'));
					break;
				case 'lastYear':
					$startDateTime = date('Y-01-01 00:00:00', strtotime('first day of last year'));
					$endDateTime = date('Y-12-31 23:59:59', strtotime('last day of last year'));
					break;
				default:
					return array(
						'Status' => 'Error',
						'Message' => 'Invalid Granularity'
					);
			}
		}

	    foreach ($results as $row) {
	        $recordStart = max(strtotime($row['date_start']), strtotime($startDateTime));
	        $recordEnd = min(strtotime($row['date_end']), strtotime($endDateTime));
	        $overlapHours = max(0, ($recordEnd - $recordStart) / 3600);

	        if ($overlapHours <= 0) continue;

			$recordCount++;

	        foreach ($row as $key => $value) {
	            if (in_array($key, ['id', 'date_start', 'date_end'])) continue;
	            if (!isset($sums[$key])) $sums[$key] = 0;
	            $sums[$key] += $value * $overlapHours;
	        }
	    }

		$averages = [];
		$averages['recordCount'] = $recordCount;
		
		foreach ($sums as $key => $sum) {
			$averages[$key] = $recordCount > 0 ? $sum / $recordCount : 0;
		}

	    return $averages;
	}

}