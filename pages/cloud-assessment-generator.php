<?php
  $ibPlugin = new TemplateConfig();
  if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CLOUDASSESSMENT'] ?? null) == false) {
    $ibPlugin->api->setAPIResponse('Error','Unauthorized',401);
    return false;
  };

  $TemplateSelection = "";
  $ActiveTemplates = $ibPlugin->getCloudAssessmentActiveTemplate();
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
              <h4>Cloud Assessment Report Generator</h4>
              <p>You can use this tool to generate Cloud Assessment Reports for Infoblox Portal accounts.</p>
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
                  <input class="form-control APIKey" onkeyup="checkInput(this.value)" id="CAGAPIKey" type="password" placeholder="Enter API Key" required>
                  <i class="fas fa-save saveBtn" onclick="apiKeyBtn(this);"></i>
              </div>
              <div class="col-md-2 realm">
                  <select id="CAGRealm" class="form-select" aria-label="Realm Selection">
                      <option value="US" selected>US Realm</option>
                      <option value="EU">EU Realm</option>
                  </select>
              </div>
              <div class="col-auto actions">
                <button class="btn btn-success" id="GenerateCA">Generate</button>
              </div>
            </div>
            <div class="row justify-content-md-center toolsMenu pt-2">
              <div class="col-lg-8">
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
                              <select  id="CAGtemplateSelection" class="select2" name="templateSelection" multiple="multiple">
                                $TemplateSelection
                              </select>
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
    <div class="row mx-2 cag-loading-div" style="display: none;">
      <div class="col-lg-12">
        <div class="card">
          <div class="card-body">
            <div class="cag-loading-icon">
              <br>
              <div class="alert alert-info genInfo" role="alert">
                <center>It can take up to 3 minutes to generate the report(s), please be patient.</center>
              </div>
              <hr>
              <div class="progress">
                <div id="cag-progress-bar" class="progress-bar" role="progressbar" style="width: 0%;" aria-valuenow="0" aria-valuemin="0" aria-valuemax="100">0%</div>
              </div>
              <br>
              <div id="spinner-container">
                <div class="spinner-bounce">
                  <div class="spinner-child spinner-bounce1"></div>
                  <div class="spinner-child spinner-bounce2"></div>
                  <div class="spinner-child spinner-bounce3"></div>
                </div>
              </div>
              <p class="progressAction" id="cag-progressAction"></p>
              <small id="cag-elapsed"></small>
            </div>
          </div>
        </div>
      </div>
    </div>
  </section>
  
  
  <script>
  var haltCAGProgress = false;
  
  function download(url) {
    const a = document.createElement("a")
    a.href = url
    a.download = url.split("/").pop()
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
  }
  
  function showCAGLoading(id,timer) {
    document.querySelector(".cag-loading-icon").style.display = "block";
    document.querySelector(".cag-loading-div").style.display = "block";
    haltCAGProgress = false;
    updateCAGProgress(id,timer);
  }
  
  function hideCAGLoading(timer) {
    document.querySelector(".cag-loading-icon").style.display = "none";
    document.querySelector(".cag-loading-div").style.display = "none";
    haltCAGProgress = true;
    stopTimer(timer);
  }
  
  function updateCAGProgress(id,timer) {
    queryAPI("GET", "/api/plugin/ib/assessment/cloud/progress?id="+id).done(function(response) {
        var data = response.data;
        var progress = parseFloat(data["Progress"]).toFixed(1); // Assuming the server returns a JSON object with a "progress" field
        $("#cag-progress-bar").css("width", progress + "%").attr("aria-valuenow", progress).text(progress + "%");
        $("#cag-progressAction").text(data["Action"])
        if (progress < 100 && haltCAGProgress == false) {
          setTimeout(function() {
            updateCAGProgress(id,timer);
          }, 1000);
        } else if (progress >= 100 && data["Action"] == "Done.." ) {
          toast("Success","","Cloud Assessment Successfully Generated","success","5000");
          toast("Success","Please wait..","Downloading Report(s)..","info","30000");
          download("/api/plugin/ib/assessment/cloud/download?id="+id);
          hideCAGLoading(timer);
          $("#GenerateCA").prop("disabled", false);
        }
    }).fail(function( data, status ) {
      setTimeout(function() {
        updateCAGProgress(id,timer);
      }, 1000);
    });
  }
  
  $("#changelog-modal-button").click(function(){
    $("#changelog-modal").modal("show")
  })
  
  $("#GenerateCA").click(function(){
    if (!$("#CAGAPIKey").is(":disabled")) {
      if(!$("#CAGAPIKey")[0].value) {
      toast("Error","Missing Required Fields","The API Key is a required field.","danger","30000");
      return null;
      }
    }
  
    $("#GenerateCA").prop("disabled", true)
    queryAPI("GET", "/api/uuid/generate").done(function( data ) {
      if (data.data) {
        let id = data.data;
        let CAGtimer = startTimer('#cag-elapsed');
        showCAGLoading(id,CAGtimer);
        var postArr = {};
        postArr.Realm = $("#CAGRealm").find(":selected").val();
        postArr.id = id;
        postArr.templates = $("#CAGtemplateSelection").val();

        if ($("#CAGAPIKey")[0].value) {
          postArr.APIKey = $("#CAGAPIKey")[0].value;
        }
        queryAPI("POST","/api/plugin/ib/assessment/cloud/generate", postArr).done(function( data, status ) {
          if (data["result"] == "Success") {
            toast("Success","Do not refresh the page","Cloud Assessment Report Job Started Successfully","success","30000");
          } else {
            toast(data["result"],"",data["message"],"danger","30000");
            hideCAGLoading(CAGtimer);
            $("#GenerateCA").prop("disabled", false);
          }
        }).fail(function( data, status ) {
            toast("API Error","","Unknown API Error","danger","30000");
            hideCAGLoading(CAGtimer);
            $("#GenerateCA").prop("disabled", false);
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