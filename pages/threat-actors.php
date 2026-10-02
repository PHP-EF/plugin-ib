<?php
  $ibPlugin = new ibPlugin();
  if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-THREATACTORS'] ?? null) == false) {
    $ibPlugin->api->setAPIResponse('Error','Unauthorized',401);
    return false;
  };
  return <<<EOF
  <section class="section">
    <div class="row mx-2">
      <div class="card">
        <div class="card-body">
          <center>
            <h4>Threat Actor Query</h4>
            <p>You can use this tool to perform queries on Threat Actors found in a particular Infoblox Portal account.</p>
          </center>
        </div>
      </div>
    </div>
    <br>
    <div class="row mx-2">
      <div class="card">
        <div class="card-body">
          <div class="container">
            <div class="row justify-content-md-center toolsMenu">
              <div class="col-md-4 apiKey">
                  <input class="form-control APIKey" onkeyup="checkInput(this.value)" id="TARAPIKey" type="password" placeholder="Enter API Key" required>
                  <i class="fas fa-save saveBtn" onclick="apiKeyBtn(this);"></i>
              </div>
              <div class="col-md-2 realm">
                  <select id="TARRealm" class="form-select" aria-label="Realm Selection">
                      <option value="US" selected>US Realm</option>
                      <option value="EU">EU Realm</option>
                  </select>
              </div>
              <div class="col-md-3">
                  <input type="text" id="TARassessmentStartAndEndDate" class="assessmentStartAndEndDate" placeholder="Start & End Date/Time">
              </div>
              <div class="col-auto actions">
                <button class="btn btn-success" id="Actors">Generate</button>
              </div>
            </div>
            <div class="row justify-content-md-center toolsMenu pt-2">
              <div class="col-lg-9">
                <div class="accordion" id="threatActorOptionsAccordion">
                  <div class="accordion-item">
                    <h2 class="accordion-header" id="threatActorOptionsHeading">
                      <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#threatActorOptions" aria-expanded="true" aria-controls="threatActorOptions">
                      Threat Actor Options
                      </button>
                    </h2>
                    <div id="threatActorOptions" class="accordion-collapse collapse" aria-labelledby="threatActorOptionsHeading" data-bs-parent="#threatActorOptionsAccordion">
                      <div class="accordion-body">
                        <div class="card-body" id="threatActorOptionsCard">
                          <div class="row">
                            <div class="col-md-6">
                              <div class="row">
                                <div class="col-md-6">
                                  <div class="form-group">
                                    <div class="form-check form-switch">
                                      <input class="form-check-input" type="checkbox" id="TARunnamed" name="unnamed">
                                      <label class="form-check-label" for="unnamed">Enable Unnamed Actors</label>
                                    </div>
                                  </div>
                                </div>
                                <div class="col-md-6">
                                  <div class="form-group">
                                    <div class="form-check form-switch">
                                      <input class="form-check-input" type="checkbox" id="TARsubstring" name="substring">
                                      <label class="form-check-label" for="substring">Enable Substring_* Actors</label>
                                    </div>
                                  </div>
                                </div>
                                <div class="col-md-6">
                                  <div class="form-group">
                                    <div class="form-check form-switch">
                                      <input class="form-check-input" type="checkbox" id="TARunknown" name="unknown">
                                      <label class="form-check-label" for="unknown">Enable Unknown Actors</label>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
            <br>
          </div>
        </div>
      </div>
    </div>
    <br>
    <div class="row mx-2 tar-loading-div" style="display: none;">
      <div class="card">
        <div class="card-body">
          <div class="tar-loading-icon">
            <div class="alert alert-info genInfo" role="alert">
              <center>It can take up to 2 minutes to generate the list of Threat Actors, please be patient.</center>
            </div>
            <br>
            <div id="spinner-container">
              <div class="spinner-bounce">
                <div class="spinner-child spinner-bounce1"></div>
                <div class="spinner-child spinner-bounce2"></div>
                <div class="spinner-child spinner-bounce3"></div>
              </div>
            </div>
            <small id="elapsed"></small>
          </div>
          <table id="threatActorTable" class="table table-striped rounded"></table>
        </div>
      </div>
    </div>
  </section>

  <!-- Observed IOC Modal -->
  <div class="modal fade" id="observedIOCModal" tabindex="-1" role="dialog" aria-labelledby="observedIOCModalLabel" aria-hidden="true">
      <div class="modal-dialog modal-xl" role="document">
          <div class="modal-content">
              <div class="modal-header">
                  <h5 class="modal-title" id="observedIOCModalLabel">Observed Indicators</h5>
                  <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close">
                      <span aria-hidden="true"></span>
                  </button>
              </div>
              <div class="modal-body" id="observedIOCModalBody">
                  <table id="threatActorObservedIOCTable" class="table table-striped rounded"></table>
              </div>
              <div class="modal-footer">
                  <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
              </div>
          </div>
      </div>
  </div>

  <!-- Related IOC Modal -->
  <div class="modal fade" id="relatedIOCModal" tabindex="-1" role="dialog" aria-labelledby="relatedIOCModalLabel" aria-hidden="true">
      <div class="modal-dialog modal-xl" role="document">
          <div class="modal-content">
              <div class="modal-header">
                  <h5 class="modal-title" id="relatedIOCModalLabel">Related Indicators</h5>
                  <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close">
                      <span aria-hidden="true"></span>
                  </button>
              </div>
              <div class="modal-body" id="relatedIOCModalBody">
                  <input id="threatActorID" hidden></input>
                  <table id="threatActorRelatedIOCTable" class="table table-striped rounded"></table>
              </div>
              <div class="modal-footer">
                  <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
              </div>
          </div>
      </div>
  </div>

  <script>
    var maxDaysApart = 31;
    var today = new Date();
    var maxPastDate = new Date(today);
    maxPastDate.setDate(today.getDate() - 31);

    flatpickr("#TARassessmentStartAndEndDate", {
      mode: "range",
      minDate: maxPastDate,
      maxDate: today,
      enableTime: true,
      dateFormat: "Y-m-d H:i",
      onChange: function(selectedDates, dateStr, instance) {
        if (selectedDates.length === 1) {
          const startDate = selectedDates[0];
          const maxEndDate = new Date(startDate.getTime() + 31 * 24 * 60 * 60 * 1000); // 31 days later
          const today = new Date();
          instance.set('maxDate', maxEndDate > today ? today : maxEndDate);
        }
        if (selectedDates.length === 2) {
          const startDate = selectedDates[0];
          const endDate = selectedDates[1];
          const diffInDays = (endDate - startDate) / (1000 * 60 * 60 * 24);
          if (diffInDays > 31) {
            toast("Error","","The start and end date cannot exceed 31 days.","warning");
            instance.clear();
          }
        }
      }
    });

    function showTARLoading(TARtimer) {
      $("#GetActors").prop("disabled", true)
      document.querySelector(".tar-loading-icon").style.display = "block";
      document.querySelector(".tar-loading-div").style.display = "block";
    }

    function hideTARLoading(TARtimer) {
      $("#GetActors").prop("disabled", false)
      document.querySelector(".tar-loading-icon").style.display = "none";
      $("#progress-bar").css("width", "0%").attr("aria-valuenow", 0).text("0%");
      stopTimer(TARtimer);
    }

    function iocCountFormatter(value, row, index) {
      if (value) {
          return value.length;
      }
    }

    function actionFormatter(value, row, index) {
      return [
        `<a class="observed" title="Observed IOCs" style="padding:5px">`,
        `<i class="fa-regular fa-eye"></i>`,
        `</a>`,
        `<a class="related" title="Related IOCs" style="padding:5px">`,
        `<i class="fa-solid fa-bullseye"></i>`,
        "</a>"
      ].join("")
    }

    function dateFormatter(value, row, index) {
      var d = new Date(value) // The 0 there is the key, which sets the date to the epoch
      return d.toGMTString();
    }

    function populateRelatedIOCs(row) {
      // var related_indicators = []
      // row["related_indicators"].forEach(function(val) {
      //   related_indicators.push({"ioc":val})
      // })
      $("#threatActorID").val(row.actor_id);
      $("#threatActorRelatedIOCTable").bootstrapTable("destroy");
      $("#threatActorRelatedIOCTable").bootstrapTable({
        ajax: "getThreatActor",
        sortable: true,
        pagination: true,
        sidePagination: "server",
        pageSize: "10000",
        queryParamsType: "",
        paginationVAlign: "both",
        showExport: true,
        exportTypes: ["json", "xml", "csv", "txt", "excel", "sql"],
        showColumns: true,
        paginationParts: ["pageInfo","pageList"],
        columns: [{
          field: "ioc",
          title: "Indicator",
          sortable: true
        }]
      });
    }
    function getThreatActor(params) {
      var postArr = {}
      postArr.Realm = $("#TARRealm").find(":selected").val();
      postArr.Page = params.data.pageNumber || 1;
      if ($("#TARAPIKey")[0].value) {
        postArr.APIKey = $("#TARAPIKey")[0].value
      }
      queryAPI("POST", "/api/plugin/ib/threatactor/"+$("#threatActorID").val(), postArr).done(function( data, status ) {
        const mappedObject = {
          total: data["data"]["actors"][0]["related_count"],
          rows: data["data"]["actors"][0]["related_indicators"].map(indicator => ({ ioc: indicator }))
        };
        params.success(mappedObject);
      }).fail(function( data, status ) {
          toast("API Error","","Failed to query IOCs","danger","30000");
      })
    }

    // Workaround
    function populateObservedIOCs(row) {
      $("#threatActorObservedIOCTable").bootstrapTable("destroy");
      $("#threatActorObservedIOCTable").bootstrapTable({
        data: row["observed_iocs"],
        sortable: true,
        pagination: true,
        search: true,
        showExport: true,
        exportTypes: ["json", "xml", "csv", "txt", "excel", "sql"],
        showColumns: true,
        columns: [{
          field: "ThreatActors.domain",
          title: "Indicator",
          sortable: true
        },{
          field: "ThreatActors.ikbfirstsubmittedts",
          title: "Submitted",
          sortable: true,
          formatter: "dateFormatter"
        },{
          field: "ThreatActors.lastdetectedts",
          title: "Last Detected",
          sortable: true,
          formatter: "dateFormatter"
        },{
          field: "ThreatActors.vtfirstdetectedts",
          title: "Virus Total Detected",
          sortable: true,
          formatter: "dateFormatter"
        }]
      });
    }

    window.actionEvents = {
      "click .observed": function (e, value, row, index) {
        populateObservedIOCs(row);
        $("#observedIOCModal").modal("show");
      },
      "click .related": function (e, value, row, index) {
        populateRelatedIOCs(row);
        $("#relatedIOCModal").modal("show");
      }
    }

    $("#Actors").on("click",function(e) {
      if (!$("#TARAPIKey").is(":disabled")) {
          if(!$("#TARAPIKey")[0].value) {
              toast("Error","Missing Required Fields","The API Key is a required field.","danger","30000");
              return null;
          }
      }
      if(!$("#TARassessmentStartAndEndDate")[0].value){
          toast("Error","Missing Required Fields","The Start Date is a required field.","danger","30000");
          return null;
      }

      let TARtimer = startTimer();
      showTARLoading(TARtimer);
      const assessmentStartAndEndDate = $("#TARassessmentStartAndEndDate")[0].value.split(" to ");
      const startDateTime = new Date(assessmentStartAndEndDate[0]);
      const endDateTime = new Date(assessmentStartAndEndDate[1]);
      var postArr = {}
      postArr.StartDateTime = startDateTime.toISOString()
      postArr.EndDateTime = endDateTime.toISOString()
      postArr.Realm = $("#TARRealm").find(":selected").val()
      postArr.unnamed = $("#TARunnamed")[0].checked;
      postArr.substring = $("#TARsubstring")[0].checked;
      postArr.unknown = $("#TARunknown")[0].checked;
      if ($("#TARAPIKey")[0].value) {
        postArr.APIKey = $("#TARAPIKey")[0].value
      }
      queryAPI("POST", "/api/plugin/ib/threatactors", postArr).done(function( data, status ) {
        if (data["result"] == "Error") {
          toast(data["result"],"",data["message"],"danger","30000");
          hideTARLoading(TARtimer);
        } else {
          $("#threatActorTable").bootstrapTable("destroy");
          $("#threatActorTable").bootstrapTable({
            data: data,
            dataField: "data",
            sortable: true,
            pagination: true,
            search: true,
            showExport: true,
            exportTypes: ["json", "xml", "csv", "txt", "excel", "sql"],
            showColumns: true,
            columns: [{
              field: "actor_name",
              title: "Name",
              sortable: true
            },{
              field: "actor_description",
              title: "Description",
              sortable: true
            },
            {
              field: "observed_iocs",
              title: "Observed IOCs",
              formatter: "iocCountFormatter",
              sortable: true
            },{
              field: "related_count",
              title: "Related IOCs",
              sortable: true
            },{
              field: "actor_id",
              title: "ID",
              sortable: false
            },{
              title: "Actions",
              formatter: "actionFormatter",
              events: "actionEvents",
            }]
          });
          hideTARLoading(TARtimer);
          return false;
        }
      }).fail(function( data, status ) {
        toast("API Error","","Unknown API Error","danger","30000");
      }).always(function() {
        hideTARLoading(TARtimer);
      });
    });
  </script>
EOF;