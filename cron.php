<?php
$scheduler->call(function() {
    try {
        $ibPlugin = new SecurityAssessment();
        $reportFiles = $ibPlugin->getReportFiles();
        $hoursBeforeExpiry = 4;
        
        $filesCleaned = array(
            'filesCleaned' => [],
            'directoriesCleaned' => []
        );
        
        foreach ($reportFiles as $reportFile) {
            $FullPath = $ibPlugin->getDir()['Files'].'/reports/'.$reportFile;
            $fileAge = time() - filemtime($FullPath);
            if ($fileAge > 4 * 3600) { // 4 hours in seconds
                if (is_file($FullPath)) {
                    if (!unlink($FullPath)) {
                        $ibPlugin->logging->writeLog("ReportCleanup","Error! Unable to delete report: ".$reportFile,"error");
                    } else {
                        $filesCleaned['filesCleaned'][] = $reportFile;
                    }
                } else {
                    if (!rmdirRecursive($FullPath)) {
                        $ibPlugin->logging->writeLog("ReportCleanup","Error! Unable to delete directory: ".$reportFile,"error");
                    } else {
                        $filesCleaned['directoriesCleaned'][] = $reportFile;
                    }
                }
            }
        }
        
        if (isset($filesCleaned['filesCleaned']) && count($filesCleaned['filesCleaned']) > 0 || isset($filesCleaned['directoriesCleaned']) && count($filesCleaned['directoriesCleaned']) > 0) {
            $ibPlugin->updateCronStatus('IB-Tools','Report Cleanup', 'success', 'Successfully cleaned up old reports');
            $ibPlugin->logging->writeLog("ReportCleanup","Successfully cleaned up old reports.","info",$filesCleaned);
        } else {
            $ibPlugin->logging->writeLog("ReportCleanup","Nothing to clean up.","info");
            $ibPlugin->updateCronStatus('IB-Tools','Report Cleanup', 'success', 'Nothing to clean up');
        }
    } catch (Exception $e) {
        $ibPlugin->updateCronStatus('IB-Tools','Report Cleanup', 'error', $e->getMessage());
    }
})->at('*/30 * * * *');