<?php
  $ibPlugin = new ibPlugin();
  if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-REPORTING'] ?: 'ACL-REPORTING') == false) {
    $ibPlugin->api->setAPIResponse('Error','Unauthorized',401);
    return false;
  };
  return <<<EOF
  <main id="main" class="main">
    <section class="section reporting-section">
      <div class="row">
        <!-- Columns -->
        <div class="col-lg-12">
          <div class="row">
            <!-- Granularity Card -->
            <div class="col-lg-3 col-md-4 col-sm-6 col-12">
              <div class="card info-card granularity-card">
                <div class="card-body">
                  <h5 class="card-title">Granularity</span></h5>
                  <div class="d-flex align-items-center">
                    <div class="btn-group">
                      <button id="granularityBtn" class="btn btn-secondary btn-sm dropdown-toggle" data-granularity="last30Days" type="button" data-bs-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                        Last 30 Days
                      </button>
                      <div class="dropdown-menu">
                        <a class="dropdown-item anon-granularity-select preventDefault" data-granularity="today" href="#">Today</a>
                        <a class="dropdown-item anon-granularity-select preventDefault" data-granularity="last30Days" href="#">Last 30 Days</a>
                        <a class="dropdown-item anon-granularity-select preventDefault" data-granularity="thisWeek" href="#">This Week</a>
                        <a class="dropdown-item anon-granularity-select preventDefault" data-granularity="thisMonth" href="#">This Month</a>
                        <a class="dropdown-item anon-granularity-select preventDefault" data-granularity="thisYear" href="#">This Year</a>
                        <a class="dropdown-item anon-granularity-select preventDefault" data-granularity="lastMonth" href="#">Last Month</a>
                        <a class="dropdown-item anon-granularity-select preventDefault" data-granularity="lastYear" href="#">Last Year</a>
                        <a class="dropdown-item anon-granularity-select preventDefault" data-granularity="custom" href="#">Custom</a>
                      </div>
                      <button id="clearFilters" class="btn btn-info btn-sm clearFilters" type="button">
                        Clear Filters
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            </div><!-- Granularity Card -->
          </div>

          <div class="row" id="metrics-container"></div>

        </div><!-- End columns -->
      </div>
    </section>

  </main><!-- End #main -->

  <!-- Custom date range modal -->
  <div class="modal fade" id="anonCustomDateRangeModal" tabindex="-1" role="dialog" aria-labelledby="anonCustomDateRangeModalLabel" aria-hidden="true">
    <div class="modal-dialog" role="document">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="anonCustomDateRangeModalLabel">Select Custom Date Range</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close">
            <span aria-hidden="true"></span>
          </button>
        </div>
        <div class="modal-body">
          <div class="toolsMenu">
            <label for="anonReportingStartAndEndDate">Select Date and Time Range:</label>
            <div class="col-md-12">
              <input type="text" id="anonReportingStartAndEndDate" placeholder="Start & End Date/Time">
            </div>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
          <button type="button" id="anonApplyCustomRange" class="btn btn-primary">Apply</button>
        </div>
      </div>
    </div>
  </div>

  <script>

    loadAnonymisedMetrics = (granularity,start, end) => {
      // Clear existing metrics
      const container = document.getElementById("metrics-container");
      container.innerHTML = '';

      // Fetch data from the API
      queryAPI("GET", "/api/plugin/ib/anonymised/security/averages?granularity="+granularity+"&start="+start+"&end="+end).done(function( response, status ) {
        var data = response['data'];
        
        // Friendly name mapping
        const friendlyNames = {
          recordCount: "Cloud Assessment Count",
          dns_requests: "DNS Requests",
          security_events_high_risk: "High Risk Events",
          security_events_medium_risk: "Medium Risk Events",
          security_events_low_risk: "Low Risk Events",
          security_events_doh: "DNS Over HTTPS Events",
          security_events_zero_day: "Zero Day DNS Events",
          security_events_suspicious: "Suspicious DNS Events",
          security_events_newly_observed_domains: "Newly Observed Domains",
          security_events_dga: "DGA Events",
          security_events_tunnelling: "Tunnelling Events",
          security_insights: "Security Insights",
          security_threat_actors: "Threat Actors",
          web_unique_applications: "Unique Web Applications",
          web_high_risk_categories: "High Risk Web Categories",
          lookalikes_custom_domains: "Lookalike Custom Domains"
        };

        const container = document.getElementById("metrics-container");

        Object.entries(data).forEach(([key, value]) => {
          const title = friendlyNames[key] || key.replace(/_/g, " ");
          const formattedValue = Number(value).toLocaleString(undefined, { maximumFractionDigits: 2 });

          const card = document.createElement("div");
          card.className = "col-lg-auto col-md-4 col-sm-6 col-12";
          card.innerHTML = '<div class="card info-card reports-'+granularity.toLowerCase()+'-card"><div class="card-body"><h5 class="card-title">'+title+' <span class="granularity-title">| '+granularity+'</span></h5><div class="d-flex align-items-center"><div class="pt-1 ps-3"><h6 class="metric-circle border-5">'+formattedValue+'</h6></div></div></div></div>'
          container.appendChild(card);
          console.log(title, formattedValue);
        });

        $(".granularity-title").text(" | "+$("#granularityBtn").text());
      });
    };

    // document.addEventListener("DOMContentLoaded", () => {
    // Declare a global variable to store active filters
    var appliedFilters = {};
    function resetAppliedFilters() {
      appliedFilters = {
        placeholder: "all"
      };
      $("#clearFilters").css("display","none");
    }

    $("#anonApplyCustomRange").on("click", function(event) {
      chartTimeFilter();
      $("#anonCustomDateRangeModal").modal("hide");
    });

    // Filter the chart with custom date/time range
    function chartTimeFilter() {
      if($("#granularityBtn").attr("data-granularity") == "custom") {
        if(!$("#anonReportingStartAndEndDate")[0].value){
          toast("Error","Missing Required Fields","The Start & End Date is a required field.","danger","30000");
          return null;
        }
        const anonReportingStartAndEndDate = $("#anonReportingStartAndEndDate")[0].value.split(" to ");
        const startDateTime = (new Date(anonReportingStartAndEndDate[0])).toISOString();
        const endDateTime = (new Date(anonReportingStartAndEndDate[1])).toISOString();
        loadAnonymisedMetrics($("#granularityBtn").attr("data-granularity"),startDateTime,endDateTime);
      } else {
        loadAnonymisedMetrics($("#granularityBtn").attr("data-granularity"));
      }
    }

    // Granularity Button
    $(".anon-granularity-select").on("click", function(event) {
      if ($(event.currentTarget).data("granularity") == "custom") {
        $("#anonCustomDateRangeModal").modal("show");
      } else {
        loadAnonymisedMetrics($(event.currentTarget).data("granularity"));
      }
      $("#granularityBtn").text($(event.currentTarget).text()).attr("data-granularity",$(event.currentTarget).data("granularity"));
    });

    // Filter Button
    $("#clearFilters").on("click", function(event) {
      // Reset Applied Filters
      resetAppliedFilters();
    })

    // Initial render
    resetAppliedFilters();
    loadAnonymisedMetrics("last30Days");
  </script>
EOF;