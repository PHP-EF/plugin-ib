<?php
class AssessmentReporting extends ibPlugin {
    public function __construct() {
		parent::__construct();
        // Create or open the SQLite database
        $this->createReportingTables();
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
}