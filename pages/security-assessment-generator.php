<?php
  $ibPlugin = new TemplateConfig();
  if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-SECURITYASSESSMENT'] ?? null) == false) {
    $ibPlugin->api->setAPIResponse('Error','Unauthorized',401);
    return false;
  };

  $TemplateSelection = "";
  $ActiveTemplates = $ibPlugin->getSecurityAssessmentActiveTemplate();
  if (is_array($ActiveTemplates) && count($ActiveTemplates) > 1) {
    foreach ($ActiveTemplates as $activeTemplate) {
      if ($activeTemplate['isDefault'] == 'true') {
        $Selected = 'selected';
      } else {
        $Selected = '';
      }
      $TemplateSelection .= '<option value="'.$activeTemplate['id'].'" '.$Selected.'>'.$activeTemplate['TemplateName'].'</option>';
    }
  } else {
    if ($ActiveTemplates[0]['isDefault'] == 'true') {
      $Selected = 'selected';
    } else {
      $Selected = '';
    }
    $TemplateSelection .= '<option value="'.$ActiveTemplates[0]['id'].'" '.$Selected.'>'.$ActiveTemplates[0]['TemplateName'].'</option>';
  }

  return <<<EOF
  <section class="section">
    <div class="row mx-2">
      <div class="col-lg-12">
        <div class="card">
          <div class="card-body">
            <center>
              <h4>Security Assessment Report Generator</h4>
              <p>You can use this tool to generate Security Assessment Reports for Infoblox Portal accounts.</p>
            </center>
          </div>
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
                  <input class="form-control APIKey" onkeyup="checkInput(this.value)" id="SAGAPIKey" type="password" placeholder="Enter API Key" required>
                  <i class="fas fa-save saveBtn" onclick="apiKeyBtn(this);"></i>
              </div>
              <div class="col-md-2 realm">
                  <select id="SAGRealm" class="form-select" aria-label="Realm Selection">
                      <option value="US" selected>US Realm</option>
                      <option value="EU">EU Realm</option>
                  </select>
              </div>
              <div class="col-md-3">
                  <input type="text" id="SAGassessmentStartAndEndDate" class="assessmentStartAndEndDate" placeholder="Start & End Date/Time">
              </div>
              <div class="col-auto actions">
                <button class="btn btn-success" id="GenerateSA">Generate</button>
              </div>
            </div>
            <div class="row justify-content-md-center toolsMenu pt-2">
              <div class="col-lg-9">
                <div class="accordion" id="assessmentOptionsAccordion">
                  <div class="accordion-item">
                    <h2 class="accordion-header" id="assessmentOptionsHeading">
                      <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#assessmentOptions" aria-expanded="true" aria-controls="assessmentOptions">
                      Assessment Options
                      </button>
                    </h2>
                    <div id="assessmentOptions" class="accordion-collapse collapse" aria-labelledby="assessmentOptionsHeading" data-bs-parent="#assessmentOptionsAccordion">
                      <div class="accordion-body">
                        <div class="card-body" id="assessmentOptionsCard">
                          <div class="row">
                            <div class="col-md-6">
                              <label class="form-check-label" for="templateSelection">Template(s)</label>
                              <select  id="SAGtemplateSelection" class="select2" name="templateSelection" multiple="multiple">
                                $TemplateSelection
                              </select>
                            </div>
                            <div class="col-md-6">
                              <div class="row">
                                <label>Threat Actor Options</label>
                                <div class="col-md-6">
                                  <div class="form-group">
                                    <div class="form-check form-switch">
                                      <input class="form-check-input" type="checkbox" id="SAGallTAInMetrics" name="allTAInMetrics" checked>
                                      <label class="form-check-label" for="allTAInMetrics">Include All Threat Actors in Count</label>
                                    </div>
                                  </div>
                                </div>
                                <div class="col-md-6">
                                  <div class="form-group">
                                    <div class="form-check form-switch">
                                      <input class="form-check-input" type="checkbox" id="SAGunnamed" name="unnamed">
                                      <label class="form-check-label" for="unnamed">Enable Unnamed Actors</label>
                                    </div>
                                  </div>
                                </div>
                                <div class="col-md-6">
                                  <div class="form-group">
                                    <div class="form-check form-switch">
                                      <input class="form-check-input" type="checkbox" id="SAGsubstring" name="substring">
                                      <label class="form-check-label" for="substring">Enable Substring_* Actors</label>
                                    </div>
                                  </div>
                                </div>
                                <div class="col-md-6">
                                  <div class="form-group">
                                    <div class="form-check form-switch">
                                      <input class="form-check-input" type="checkbox" id="SAGunknown" name="unknown">
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
    <div class="row mx-2 sag-loading-div" style="display: none;">
      <div class="col-lg-12">
        <div class="card">
          <div class="card-body">
            <div class="sag-loading-icon">
              <br>
              <div class="alert alert-info genInfo" role="alert">
                <center>It can take up to 3 minutes to generate the report(s), please be patient.</center>
              </div>
              <hr>
              <div class="progress">
                <div id="sag-progress-bar" class="progress-bar" role="progressbar" style="width: 0%;" aria-valuenow="0" aria-valuemin="0" aria-valuemax="100">0%</div>
              </div>
              <br>
              <div id="spinner-container">
                <div class="spinner-bounce">
                  <div class="spinner-child spinner-bounce1"></div>
                  <div class="spinner-child spinner-bounce2"></div>
                  <div class="spinner-child spinner-bounce3"></div>
                </div>
              </div>
              <p class="progressAction" id="sag-progressAction"></p>
              <small id="sag-elapsed"></small>
            </div>
          </div>
        </div>
      </div>
    </div>
  </section>
  
  
  <script>
  var haltProgress = false;
  var maxDaysApart = 31;
  var today = new Date();
  var maxPastDate = new Date(today);
  maxPastDate.setDate(today.getDate() - 31);

  flatpickr("#SAGassessmentStartAndEndDate", {
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
  
  function download(url) {
    const a = document.createElement("a")
    a.href = url
    a.download = url.split("/").pop()
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
  }
  
  function showSAGLoading(id,timer) {
    document.querySelector(".sag-loading-icon").style.display = "block";
    document.querySelector(".sag-loading-div").style.display = "block";
    haltProgress = false;
    updateSAGProgress(id,timer);
  }
  
  function hideSAGLoading(timer) {
    document.querySelector(".sag-loading-icon").style.display = "none";
    document.querySelector(".sag-loading-div").style.display = "none";
    haltProgress = true;
    stopTimer(timer);
  }
  
  function updateSAGProgress(id,timer) {
    queryAPI("GET", "/api/plugin/ib/assessment/security/progress?id="+id).done(function(response) {
        var data = response.data;
        var progress = parseFloat(data["Progress"]).toFixed(1); // Assuming the server returns a JSON object with a "progress" field
        $("#sag-progress-bar").css("width", progress + "%").attr("aria-valuenow", progress).text(progress + "%");
        $("#sag-progressAction").text(data["Action"])
        if (progress < 100 && haltProgress == false) {
          setTimeout(function() {
            updateSAGProgress(id,timer);
          }, 1000);
        } else if (progress >= 100 && data["Action"] == "Done.." ) {
          toast("Success","","Security Assessment Successfully Generated","success","5000");
          toast("Success","Please wait..","Downloading Report(s)..","info","30000");
          download("/api/plugin/ib/assessment/security/download?id="+id);
          hideSAGLoading(timer);
          $("#GenerateSA").prop("disabled", false);
        }
    }).fail(function( data, status ) {
      setTimeout(function() {
        updateSAGProgress(id,timer);
      }, 1000);
    });
  }
  
  $("#changelog-modal-button").click(function(){
    $("#changelog-modal").modal("show")
  })
  
  $("#GenerateSA").click(function(){
    if (!$("#SAGAPIKey").is(":disabled")) {
      if(!$("#SAGAPIKey")[0].value) {
      toast("Error","Missing Required Fields","The API Key is a required field.","danger","30000");
      return null;
      }
    }
    if(!$("#SAGassessmentStartAndEndDate")[0].value){
      toast("Error","Missing Required Fields","The Start & End Date is a required field.","danger","30000");
      return null;
    }
  
    $("#GenerateSA").prop("disabled", true)
    queryAPI("GET", "/api/uuid/generate").done(function( data ) {
      if (data.data) {
        let id = data.data;
        let SAGtimer = startTimer('#sag-elapsed');
        showSAGLoading(id,SAGtimer);
        const assessmentStartAndEndDate = $("#SAGassessmentStartAndEndDate")[0].value.split(" to ");
        const startDateTime = new Date(assessmentStartAndEndDate[0]);
        const endDateTime = new Date(assessmentStartAndEndDate[1]);
        var postArr = {};
        postArr.StartDateTime = startDateTime.toISOString();
        postArr.EndDateTime = endDateTime.toISOString();
        postArr.Realm = $("#SAGRealm").find(":selected").val();
        postArr.id = id;
        postArr.templates = $("#SAGtemplateSelection").val();
        postArr.unnamed = $("#SAGunnamed")[0].checked;
        postArr.substring = $("#SAGsubstring")[0].checked;
        postArr.unknown = $("#SAGunknown")[0].checked;
        postArr.allTAInMetrics = $("#SAGallTAInMetrics")[0].checked;

        if ($("#SAGAPIKey")[0].value) {
          postArr.APIKey = $("#SAGAPIKey")[0].value;
        }
        queryAPI("POST","/api/plugin/ib/assessment/security/generate", postArr).done(function( data, status ) {
          if (data["result"] == "Success") {
            toast("Success","Do not refresh the page","Security Assessment Report Job Started Successfully","success","30000");
          } else {
            toast(data["result"],"",data["message"],"danger","30000");
            hideSAGLoading(SAGtimer);
            $("#GenerateSA").prop("disabled", false);
          }
        }).fail(function( data, status ) {
            toast("API Error","","Unknown API Error","danger","30000");
            hideSAGLoading(SAGtimer);
            $("#GenerateSA").prop("disabled", false);
        })
      } else {
        toast("API Error","","UUID not returned from the API","danger","30000");
      }
    });
  });
  
  // Initialise Select2 Inputs
  $('.select2').select2({tags: false, closeOnSelect: true, allowClear: true, width: "100%"});
  </script>
EOF;