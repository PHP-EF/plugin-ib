<?php
  $ibPlugin = new ibPlugin();
  if ($ibPlugin->auth->checkAccess($ibPlugin->config->get('Plugins','IB-Tools')['ACL-CONFIG'] ?: 'ACL-CONFIG') == false) {
    $ibPlugin->api->setAPIResponse('Error','Unauthorized',401);
    return false;
  };
  return <<<EOF
  <style>
  .card {
    padding: 10px;
  }
  </style>

  <div class="container">
    <div class="row justify-content-center">
      <div class="col-12 col-lg-12 mx-auto">
        <h2 class="h3 mb-4 page-title">Security Assessment Report Generator Configuration</h2>

        <div class="my-4">
          <div class="card border-secondary">
            <div class="card-title">
              <h5>Threat Actor Configuration</h5>
              <p>Use the following to configure the known Threat Actors. This allows populating Threat Actors with Images / Report Links during Security Assessment Report generation.</p>
            </div>
            <table  data-url="/api/plugin/ib/threatactors/config"
              data-data-field="data"
              data-toggle="table"
              data-search="true"
              data-filter-control="true"
              data-show-refresh="true"
              data-pagination="true"
              data-toolbar="#toolbar"
              data-sort-name="Name"
              data-sort-order="asc"
              data-page-size="25"
              data-buttons="threatActorButtons"
              data-buttons-order="btnAddThreatActor,refresh"
              class="table table-striped" id="threatActorConfigurationTable">

              <thead>
                <tr>
                  <th data-field="state" data-checkbox="true"></th>
                  <th data-field="Name" data-sortable="true">Threat Actor</th>
                  <th data-field="PNG" data-sortable="true">PNG Image</th>
                  <th data-field="SVG" data-sortable="true">SVG Image</th>
                  <th data-field="URLStub" data-sortable="true">URL Stub</th>
                  <th data-formatter="threatActorActionFormatter" data-events="actionEvents">Actions</th>
                </tr>
              </thead>
            </table>
          </div>

          <hr>

          <div class="card border-secondary">
            <div class="card-title">
              <h5>Template Configuration</h5>
              <p>Use the following to configure the template for the Security Assessment Report Generator.</p>
            </div>
            <table  data-url="/api/plugin/ib/assessment/security/config"
              data-data-field="data"
              data-toggle="table"
              data-search="true"
              data-filter-control="true"
              data-show-refresh="true"
              data-pagination="true"
              data-toolbar="#toolbar"
              data-sort-name="Status"
              data-sort-order="asc"
              data-page-size="25"
              data-buttons="templateButtons"
              data-buttons-order="btnAddTemplate,refresh"
              class="table table-striped" id="securityAssessmentTemplateTable">

              <thead>
                <tr>
                  <th data-field="state" data-checkbox="true"></th>
                  <th data-field="Status" data-sortable="true">Status</th>
                  <th data-field="TemplateName" data-sortable="true">Name</th>
                  <th data-field="Description" data-sortable="true">Description</th>
                  <th data-field="ThreatActorSlide" data-sortable="true">Threat Actor Slide</th>
                  <th data-field="SOCInsightsSlide" data-sortable="true">SOC Insights Slide</th>
                  <th data-field="FileName" data-sortable="true">File Name</th>
                  <th data-field="Created" data-sortable="true">Upload Date</th>
                  <th data-formatter="SAtemplateActionFormatter" data-events="actionEvents">Actions</th>
                </tr>
              </thead>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>


  <!-- Edit Threat Actor Modal -->
  <div class="modal fade" id="editThreatActorModal" tabindex="-1" role="dialog" aria-labelledby="editThreatActorModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg" role="document">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="editThreatActorModalLabel">Threat Actor Information</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close">
            <span aria-hidden="true"></span>
          </button>
        </div>
        <div class="modal-body" id="editThreatActorModalBody">
          <div class="form-group">
            <div class="form-group" hidden>
              <input type="text" class="form-control info-field" id="threatActorId" aria-describedby="threatActorIdHelp" name="threatActorId" disabled>
            </div>
            <label for="threatActorName">Threat Actor Name</label>
            <input type="text" class="form-control" id="threatActorName" aria-describedby="threatActorNameHelp" disabled>
            <small id="threatActorNameHelp" class="form-text text-muted">The name of the Threat Actor.</small>
          </div>
          <div class="form-group">
            <label for="threatActorURLStub">URL Stub</label>
            <input type="text" class="form-control" id="threatActorURLStub" aria-describedby="threatActorURLStubHelp">
            <small id="threatActorURLStubHelp" class="form-text text-muted">The Threat Actor Report <b>URL</b> to use when generating Security Assessment reports.</small>
          </div>
          <div class="form-group row">
            <label for="threatActorIMGSVG" class="col-form-label">SVG Image</label>
            <div class="col-sm-5">
              <input type="file" class="form-control" id="threatActorIMGSVG" accept=".svg" aria-describedby="threatActorIMGSVGHelp">
              <small id="threatActorIMGSVGHelp" class="form-text text-muted">Upload an SVG image.</small>
            </div>
            <div class="col-sm-5">
              <img id="imagePreviewSVG" src="" alt="PNG Image Preview" style="display:none; margin-top: 10px; max-width: 100%;">
            </div>
          </div>
          <div class="form-group row">
            <label for="threatActorIMGPNG" class="col-form-label">PNG Image</label>
            <div class="col-sm-5">
              <input type="file" class="form-control" id="threatActorIMGPNG" accept=".png" aria-describedby="threatActorIMGPNGHelp">
              <small id="threatActorIMGPNGHelp" class="form-text text-muted">Upload an PNG image.</small>
            </div>
            <div class="col-sm-5">
              <img id="imagePreviewPNG" src="" alt="PNG Image Preview" style="display:none; margin-top: 10px; max-width: 100%;">
            </div>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
          <button class="btn btn-primary" id="editThreatActorSubmit">Save</button>
        </div>
      </div>
    </div>
  </div>

  <!-- New Threat Actor Modal -->
  <div class="modal fade" id="newThreatActorModal" tabindex="-1" role="dialog" aria-labelledby="newThreatActorModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg" role="document">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="newThreatActorModalLabel">New Threat Actor Wizard</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close">
            <span aria-hidden="true"></span>
          </button>
        </div>
        <div class="modal-body" id="newThreatActorModelBody">
          <p>Enter the Threat Actor's Name below to add it to the list.</p>
          <div class="form-group">
            <label for="newThreatActorName">Threat Actor Name</label>
            <input type="text" class="form-control" id="newThreatActorName" aria-describedby="newThreatActorNameHelp">
            <small id="newThreatActorNameHelp" class="form-text text-muted">The name of the Threat Actor to add to the list.</small>
          </div>
          <div class="form-group">
            <label for="newThreatActorURLStub">URL Stub</label>
            <input type="text" class="form-control" id="newThreatActorURLStub" aria-describedby="newThreatActorURLStubHelp">
            <small id="newThreatActorURLStubHelp" class="form-text text-muted">The Threat Actor Report <b>URL</b> to use when generating Security Assessment reports.</small>
          </div>
          <div class="form-group row">
            <label for="newThreatActorIMGSVG" class="col-form-label">SVG Image</label>
            <div class="col-sm-5">
              <input type="file" class="form-control" id="newThreatActorIMGSVG" accept=".svg" aria-describedby="newThreatActorIMGSVGHelp">
              <small id="newThreatActorIMGSVGHelp" class="form-text text-muted">Upload an SVG image.</small>
            </div>
            <div class="col-sm-5">
              <img id="newImagePreviewSVG" src="" alt="PNG Image Preview" style="display:none; margin-top: 10px; max-width: 100%;">
            </div>
          </div>
          <div class="form-group row">
            <label for="newThreatActorIMGPNG" class="col-form-label">PNG Image</label>
            <div class="col-sm-5">
              <input type="file" class="form-control" id="newThreatActorIMGPNG" accept=".png" aria-describedby="newThreatActorIMGPNGHelp">
              <small id="newThreatActorIMGPNGHelp" class="form-text text-muted">Upload an PNG image.</small>
            </div>
            <div class="col-sm-5">
              <img id="newImagePreviewPNG" src="" alt="PNG Image Preview" style="display:none; margin-top: 10px; max-width: 100%;">
            </div>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
          <button class="btn btn-primary" id="newThreatActorSubmit">Save</button>
        </div>
      </div>
    </div>
  </div>

  <!-- -------------------- -->

  <!-- Edit Template Modal -->
  <div class="modal fade" id="editSATemplateModal" tabindex="-1" role="dialog" aria-labelledby="editSATemplateModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg" role="document">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="editSATemplateModalLabel">Template Information</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close">
            <span aria-hidden="true"></span>
          </button>
        </div>
        <div class="modal-body" id="editSATemplateModalBody">
          <div class="form-group">
            <div class="form-group" hidden>
              <input type="text" class="form-control info-field" id="SAtemplateId" aria-describedby="SAtemplateIdHelp" name="SAtemplateId" disabled>
            </div>
            <label for="SAtemplateStatus">Template Status</label>
            <select id="SAtemplateStatus" class="form-select" aria-label="Template Status" aria-describedby="SAtemplateStatusHelp">
              <option value="Active">Active</option>
              <option value="Inactive">Inactive</option>
            </select>
            <small id="SAtemplateStatusHelp" class="form-text text-muted">The current status of the template.</small>
          </div>
          <div class="form-group">
            <label for="SAtemplateName">Template Name</label>
            <input type="text" class="form-control info-field" id="SAtemplateName" aria-describedby="SAtemplateNameHelp" name="SAtemplateName">
            <small id="SAtemplateNameHelp" class="form-text text-muted">The name for the Security Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="SAtemplateDescription">Template Description</label>
            <input type="text" class="form-control info-field" id="SAtemplateDescription" aria-describedby="SAtemplateDescriptionHelp" name="SAtemplateDescription">
            <small id="SAtemplateDescriptionHelp" class="form-text text-muted">The description of the Security Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="SAtemplateThreatActorSlide">Threat Actor Slide</label>
            <input type="text" class="form-control info-field" id="SAtemplateThreatActorSlide" aria-describedby="SAtemplateThreatActorSlideHelp" name="SAtemplateThreatActorSlide">
            <small id="SAtemplateThreatActorSlideHelp" class="form-text text-muted">This is the Threat Actor template slide number.</small>
          </div>
          <div class="form-group">
            <label for="SAtemplateSOCInsightsSlide">SOC Insights Slide</label>
            <input type="text" class="form-control info-field" id="SAtemplateSOCInsightsSlide" aria-describedby="SAtemplateSOCInsightsSlideHelp" name="SAtemplateSOCInsightsSlide">
            <small id="SAtemplateSOCInsightsSlideHelp" class="form-text text-muted">This is the SOC Insights template slide number.</small>
          </div>
          <div class="form-group">
            <label for="SAtemplateOrientation" class="col-form-label">Template Orientation</label>
            <select id="SAtemplateOrientation" class="form-select" name="SAtemplateOrientation">
              <option value="Portrait" selected>Portrait</option>
              <option value="Landscape">Landscape</option>
            </select>
            <small id="SAtemplateOrientationHelp" class="form-text text-muted">The orientation of the powerpoint template.</small>
          </div>
          <div class="form-group">
            <label for="SAtemplateSelectedByDefault">Selected by Default</label>
            <div class="form-check form-switch">
              <input class="form-check-input info-field" type="checkbox" id="SAtemplateSelectedByDefault" name="SAtemplateSelectedByDefault" value="">  
            </div>
            <small id="SAtemplateSelectedByDefaultHelp" class="form-text text-muted">Enable this to select this template by default.</small>
          </div>
          <div class="form-group">
            <label for="SAtemplateMacroEnabled">Macro Enabled</label>
            <div class="form-check form-switch">
              <input class="form-check-input info-field" type="checkbox" id="SAtemplateMacroEnabled" name="SAtemplateMacroEnabled" value="">  
            </div>
            <small id="SAtemplateMacroEnabledHelp" class="form-text text-muted">Enable this to allow macros in the template.</small>
          </div>
          <div class="form-group row">
            <label for="SAtemplatePPTX" class="col-form-label">PowerPoint Template</label>
            <div class="col-sm-5">
              <input type="file" class="form-control" id="SAtemplatePPTX" accept=".pptx,.pptm" aria-describedby="SAtemplatePPTXHelp">
              <small id="SAtemplatePPTXHelp" class="form-text text-muted">Upload a PowerPoint Template.</small>
            </div>
            <div class="col-sm-5">
              <img id="imagePreviewSVG" src="" alt="PNG Image Preview" style="display:none; margin-top: 10px; max-width: 100%;">
            </div>
          </div>
          <div class="form-group">
            <label for="SAtemplateFileName">Template File Name</label>
            <input type="text" class="form-control info-field" id="SAtemplateFileName" aria-describedby="SAtemplateFileNameHelp" name="SAtemplateFileName" disabled>
            <small id="SAtemplateFileNameHelp" class="form-text text-muted">The file name for the Security Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="SAtemplateUploadDate">Upload Date</label>
            <input type="text" class="form-control" id="SAtemplateUploadDate" aria-describedby="SAtemplateUploadDateHelp" disabled>
            <small id="SAtemplateUploadDateHelp" class="form-text text-muted">The upload date of this template.</small>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
          <button class="btn btn-primary" id="editSATemplateSubmit">Save</button>
        </div>
      </div>
    </div>
  </div>

  <!-- New Template Modal -->
  <div class="modal fade" id="newSATemplateModal" tabindex="-1" role="dialog" aria-labelledby="newSATemplateModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg" role="document">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="newSATemplateModalLabel">New Template Wizard</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close">
            <span aria-hidden="true"></span>
          </button>
        </div>
        <div class="modal-body" id="newSATemplateModelBody">
          <p>Enter the new template information below to add it to the list.</p>
          <div class="form-group">
            <label for="newSATemplateStatus">Template Status</label>
            <select id="newSATemplateStatus" class="form-select" aria-label="Template Status" aria-describedby="newSATemplateStatusHelp">
              <option value="Active">Active</option>
              <option value="Inactive">Inactive</option>
            </select>
            <small id="newSATemplateStatusHelp" class="form-text text-muted">The status of the new template.</small>
          </div>
          <div class="form-group">
            <label for="newSATemplateName">Template Name</label>
            <input type="text" class="form-control" id="newSATemplateName" aria-describedby="newSATemplateNameHelp" name="newSATemplateName">
            <small id="newSATemplateNameHelp" class="form-text text-muted">The name for the Security Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="newSATemplateDescription">Template Description</label>
            <input type="text" class="form-control" id="newSATemplateDescription" aria-describedby="newSATemplateDescriptionHelp" name="newSATemplateDescription">
            <small id="newSATemplateDescriptionHelp" class="form-text text-muted">The description of the Security Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="newSATemplateOrientation" class="col-form-label">Template Orientation</label>
            <select id="newSATemplateOrientation" class="form-select" name="newSATemplateOrientation">
              <option value="Portrait" selected>Portrait</option>
              <option value="Landscape">Landscape</option>
            </select>
            <small id="newSATemplateOrientationHelp" class="form-text text-muted">The orientation of the powerpoint template.</small>
          </div>
          <div class="form-group">
            <label for="newSATemplateThreatActorSlide">Threat Actor Slide</label>
            <input type="text" class="form-control" id="newSATemplateThreatActorSlide" aria-describedby="newSATemplateThreatActorSlideHelp" name="newSATemplateThreatActorSlide">
            <small id="newSATemplateThreatActorSlideHelp" class="form-text text-muted">This is the Threat Actor template slide number.</small>
          </div>
          <div class="form-group">
            <label for="newSATemplateSOCInsightsSlide">SOC Insights Slide</label>
            <input type="text" class="form-control" id="newSATemplateSOCInsightsSlide" aria-describedby="newSATemplateSOCInsightsSlideHelp" name="newSATemplateSOCInsightsSlide">
            <small id="newSATemplateSOCInsightsSlideHelp" class="form-text text-muted">This is the SOC Insights template slide number.</small>
          </div>
          <div class="form-group">
            <label for="newSATemplateSelectedByDefault">Selected by Default</label>
            <div class="form-check form-switch">
              <input class="form-check-input info-field" type="checkbox" id="newSATemplateSelectedByDefault" name="newSATemplateSelectedByDefault" value="">  
            </div>
            <small id="newSATemplateSelectedByDefaultHelp" class="form-text text-muted">Enable this to select this template by default.</small>
          </div>
          <div class="form-group">
            <label for="newSATemplateMacroEnabled">Macro Enabled</label>
            <div class="form-check form-switch">
              <input class="form-check-input info-field" type="checkbox" id="newSATemplateMacroEnabled" name="newSATemplateMacroEnabled" value="">  
            </div>
            <small id="newSATemplateMacroEnabledHelp" class="form-text text-muted">Enable this to allow macros in the template.</small>
          </div>
          <div class="form-group row">
            <label for="newSATemplatePPTX" class="col-form-label">PowerPoint Template</label>
            <div class="col-sm-5">
              <input type="file" class="form-control" id="newSATemplatePPTX" accept=".pptx,.pptm" aria-describedby="newSATemplatePPTXHelp">
              <small id="newSATemplatePPTXHelp" class="form-text text-muted">Upload a PowerPoint Template.</small>
            </div>
            <div class="col-sm-5">
              <img id="imagePreviewSVG" src="" alt="PNG Image Preview" style="display:none; margin-top: 10px; max-width: 100%;">
            </div>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
          <button class="btn btn-primary" id="newSATemplateSubmit">Save</button>
        </div>
      </div>
    </div>
  </div>

  <script>
    function threatActorActionFormatter(value, row, index) {
      return [
        `<a class="editThreatActor" title="Edit">`,
        `<i class="fa fa-pencil"></i>`,
        `</a>&nbsp;`,
        `<a class="deleteThreatActor" title="Delete">`,
        `<i class="fa fa-trash"></i>`,
        `</a>`
      ].join("")
    }

    function SAtemplateActionFormatter(value, row, index) {
      return [
        `<a class="editSATemplate" title="Edit">`,
        `<i class="fa fa-pencil"></i>`,
        `</a>&nbsp;`,
        `<a class="deleteSATemplate" title="Delete">`,
        `<i class="fa fa-trash"></i>`,
        `</a>`
      ].join("")
    }

    function threatActorButtons() {
      return {
        btnAddThreatActor: {
          text: "Add Threat Actor",
          icon: "bi-plus-lg",
          event: function() {
            $("#newThreatActorModal").modal("show");
            $("#newThreatActorModal input").val("");
            $("#newImagePreviewSVG, #newImagePreviewPNG").attr("src","").css("display","none");
          },
          attributes: {
            title: "Add a new Threat Actor",
            style: "background-color:#4bbe40;border-color:#4bbe40;"
          }
        }
      }
    }

    function templateButtons() {
      return {
        btnAddTemplate: {
          text: "Add Security Assessment Template",
          icon: "bi-plus-lg",
          event: function() {
            $("#newSATemplateModal").modal("show");
            $("#newSATemplateModal input").val("");
          },
          attributes: {
            title: "Add a new Security Assessment Template",
            style: "background-color:#4bbe40;border-color:#4bbe40;"
          }
        }
      }
    }

    function listThreatActor(row) {
      $("#editThreatActorModal input").val("");
      $("#threatActorId").val(row["id"]);
      $("#imagePreviewSVG, #imagePreviewPNG").attr("src","").css("display","none");
      $("#threatActorName").val(row["Name"]);
      if (row["SVG"] != "") {
        imagePreview("imagePreviewSVG","/assets/images/Threat Actors/Uploads/"+row["PNG"]);
      }
      if (row["PNG"] != "") {
        imagePreview("imagePreviewPNG","/assets/images/Threat Actors/Uploads/"+row["PNG"]);
      }
      $("#threatActorURLStub").val(row["URLStub"]);
    }

    function listSATemplate(row) {
      $("#editSATemplateModal input").val("");
      $("#SAtemplateId").val(row["id"]);
      $("#SAtemplateStatus").val(row["Status"]);
      $("#SAtemplateName").val(row["TemplateName"]);
      $("#SAtemplateDescription").val(row["Description"]);
      $("#SAtemplateFileName").val(row["FileName"]);
      $("#SAtemplateThreatActorSlide").val(row["ThreatActorSlide"]);
      $("#SAtemplateSOCInsightsSlide").val(row["SOCInsightsSlide"]);
      $("#SAtemplateOrientation").val(row["Orientation"]);
      if (String(row["isDefault"]).toLowerCase() == "true") {
        $("#SAtemplateSelectedByDefault").attr("checked",true);
      } else {
        $("#SAtemplateSelectedByDefault").attr("checked",false);
      }
      if (String(row["macroEnabled"]).toLowerCase() == "true") {
        $("#SAtemplateMacroEnabled").attr("checked",true);
      } else {
        $("#SAtemplateMacroEnabled").attr("checked",false);
      }
      $("#SAtemplateUploadDate").val(row["Created"]);
    }

    $("#threatActorIMGPNG, #threatActorIMGSVG, #newThreatActorIMGPNG, #newThreatActorIMGSVG").on("change", function(event) {
      const file = event.target.files[0];
      let target = "";
      if (event.target.id == "threatActorIMGPNG") {
        target = "imagePreviewPNG";
      }
      if (event.target.id == "threatActorIMGSVG") {
        target = "imagePreviewSVG";
      }
      if (event.target.id == "newThreatActorIMGPNG") {
        target = "newImagePreviewPNG";
      }
      if (event.target.id == "newThreatActorIMGSVG") {
        target = "newImagePreviewSVG";
      }
      const preview = document.getElementById(target);
      if (file) {
          const reader = new FileReader();
          reader.onload = function(e) {
              preview.src = e.target.result;
              preview.style.display = "block"; // Show the image preview
          };
          reader.readAsDataURL(file);
      } else {
          preview.src = "";
          preview.style.display = "none"; // Hide the image preview if no file is selected
      }
    });

    window.actionEvents = {
      "click .editThreatActor": function (e, value, row, index) {
        listThreatActor(row);
        $("#editThreatActorModal").modal("show");
      },
      "click .deleteThreatActor": function (e, value, row, index) {
        if(confirm("Are you sure you want to delete "+row.Name+" from the list of Threat Actors? This is irriversible.") == true) {
          queryAPI("DELETE", "/api/plugin/ib/threatactors/config/"+row.id).done(function( data, status ) {
            if (data["result"] == "Success") {
              toast(data["result"],"",data["message"],"success");
              $("#threatActorConfigurationTable").bootstrapTable("refresh");
            } else if (data["result"] == "Error") {
              toast(data["result"],"",data["message"],"danger","30000");
            } else {
              toast("Error","","Failed to remove Threat Actor: "+row.Name,"danger","30000");
            }
          }).fail(function( data, status ) {
              toast("API Error","","Failed to remove Threat Actor: "+row.Name,"danger","30000");
          })
        }
      },
      "click .editSATemplate": function (e, value, row, index) {
        listSATemplate(row);
        $("#editSATemplateModal").modal("show");
      },
      "click .deleteSATemplate": function (e, value, row, index) {
        if(confirm("Are you sure you want to delete "+row.TemplateName+" from the list of Templates? This is irriversible.") == true) {
          queryAPI("DELETE", "/api/plugin/ib/assessment/security/config/"+row.id).done(function( data, status ) {
            if (data["result"] == "Success") {
              toast(data["result"],"",data["message"],"success");
              $("#securityAssessmentTemplateTable").bootstrapTable("refresh");
            } else if (data["result"] == "Error") {
              toast(data["result"],"",data["message"],"danger","30000");
            } else {
              toast("Error","","Failed to remove Template: "+row.TemplateName,"danger","30000");
            }
          }).fail(function( data, status ) {
              toast("API Error","","Failed to remove Template: "+row.TemplateName,"danger","30000");
          })
        }
      }
    }

    $(document).on("click", "#newThreatActorSubmit", function(event) {
      const svgFiles = $("#newThreatActorIMGSVG")[0].files;
      const pngFiles = $("#newThreatActorIMGPNG")[0].files;
      const threatActorFileName = $("#newThreatActorName").val().toLowerCase().replace(/ /g, "-");
      var postArr = {}
      if (svgFiles[0]) {
        postArr.SVG = threatActorFileName;
      }
      if (svgFiles[0]) {
        postArr.PNG = threatActorFileName;
      }
      postArr.name = encodeURIComponent($("#newThreatActorName").val())
      postArr.URLStub = encodeURIComponent($("#newThreatActorURLStub").val())
      queryAPI("POST", "/api/plugin/ib/threatactors/config", postArr).done(function( data, status ) {
        if (data["result"] == "Success") {
          toast(data["result"],"",data["message"],"success");
          $("#threatActorConfigurationTable").bootstrapTable("refresh");
          if (svgFiles.length > 0 || pngFiles.length > 0) {
            const formData = new FormData();
            if (svgFiles[0]) {
              formData.append("svgImage", svgFiles[0]);
              formData.append("svgFileName", threatActorFileName);
            }
            if (pngFiles[0]) {
              formData.append("pngImage", pngFiles[0]);
              formData.append("pngFileName", threatActorFileName);
            }
            $.ajax({
              url: "/api/plugin/ib/threatactors/config/upload",
              type: "POST",
              data: formData,
              contentType: false,
              processData: false,
              success: function(response) {
                if (response.data.Errors.length > 0) {
                  response.data.Errors.forEach(error => {
                    toast("Error","",error,"danger");
                  });
                }
                if (response.data.Items.length > 0) {
                  response.data.Items.forEach(item => {
                    toast("Success","",item,"success");
                  });
                }
              },
              error: function(jqXHR, textStatus, errorThrown) {
                toast("Error","","Error submitting images","danger");
                console.error(errorThrown);
              }
            });
          }
        } else if (data["result"] == "Error") {
          toast(data["result"],"",data["message"],"danger","30000");
        } else {
          toast("Error","","Failed to add new Threat Actor","danger","30000");
        }
      }).fail(function( data, status ) {
          toast("API Error","","Failed to add new Threat Actor","danger","30000");
      }).always(function( data, status) {
        $("#newThreatActorModal").modal("hide");
      })
    });

    $(document).on("click", "#editThreatActorSubmit", function(event) {
      const svgFiles = $("#threatActorIMGSVG")[0].files;
      const pngFiles = $("#threatActorIMGPNG")[0].files;
      const threatActorFileName = $("#threatActorName").val().toLowerCase().replace(/ /g, "-");
      const id = encodeURIComponent($("#threatActorId").val());
      var postArr = {}
      if (svgFiles[0]) {
        postArr.SVG = threatActorFileName;
      }
      if (svgFiles[0]) {
        postArr.PNG = threatActorFileName;
      }
      postArr.name = encodeURIComponent($("#threatActorName").val())
      postArr.URLStub = encodeURIComponent($("#threatActorURLStub").val())
      queryAPI("PATCH", "/api/plugin/ib/threatactors/config/"+id, postArr).done(function( data, status ) {
        if (data["result"] == "Success") {
          toast(data["result"],"",data["message"],"success");
          $("#threatActorConfigurationTable").bootstrapTable("refresh");
          if (svgFiles.length > 0 || pngFiles.length > 0) {
            const formData = new FormData();
            if (svgFiles[0]) {
              formData.append("svgImage", svgFiles[0]);
              formData.append("svgFileName", threatActorFileName);
            }
            if (pngFiles[0]) {
              formData.append("pngImage", pngFiles[0]);
              formData.append("pngFileName", threatActorFileName);
            }
            $.ajax({
              url: "/api/plugin/ib/threatactors/config/upload",
              type: "POST",
              data: formData,
              contentType: false,
              processData: false,
              success: function(response) {
                if (response.data.Errors.length > 0) {
                  response.data.Errors.forEach(error => {
                    toast("Error","",error,"danger");
                  });
                }
                if (response.data.Items.length > 0) {
                  response.data.Items.forEach(item => {
                    toast("Success","",item,"success");
                  });
                }
              },
              error: function(jqXHR, textStatus, errorThrown) {
                toast("Error","","Error submitting images","danger");
                console.error(errorThrown);
              }
            });
          }
        } else if (data["result"] == "Error") {
          toast(data["result"],"",data["message"],"danger","30000");
        } else {
          toast("Error","","Failed to update Threat Actor: "+postArr.name,"danger","30000");
        }
      }).fail(function( data, status ) {
          toast("API Error","","Failed to update Threat Actor: "+postArr.name,"danger","30000");
      }).always(function( data, status) {
        $("#editThreatActorModal").modal("hide");
      })
    });

    $(document).on("click", "#newSATemplateSubmit", function(event) {
      const templateFiles = $("#newSATemplatePPTX")[0].files;

      var postArr = {}
      postArr.Status = encodeURIComponent($("#newSATemplateStatus").val());
      postArr.TemplateName = encodeURIComponent($("#newSATemplateName").val());
      postArr.Description = encodeURIComponent($("#newSATemplateDescription").val());
      postArr.Orientation = encodeURIComponent($("#newSATemplateOrientation").val());
      postArr.isDefault = encodeURIComponent($("#SAtemplateSelectedByDefault")[0].checked);
      postArr.macroEnabled = encodeURIComponent($("#SAtemplateMacroEnabled")[0].checked);
      postArr.ThreatActorSlide = encodeURIComponent($("#newSATemplateThreatActorSlide").val());
      postArr.SOCInsightsSlide = encodeURIComponent($("#newSATemplateSOCInsightsSlide").val());
      if (templateFiles[0]) {
        postArr.FileName = $("#newSATemplateName").val().toLowerCase().replace(/ /g, "-");
      }
      queryAPI("POST", "/api/plugin/ib/assessment/security/config", postArr).done(function( data, status ) {
        if (data["result"] == "Success") {
          toast(data["result"],"",data["message"],"success");
          $("#securityAssessmentTemplateTable").bootstrapTable("refresh");
          if (templateFiles.length > 0) {
            const formData = new FormData();
            formData.append("pptx", templateFiles[0]);
            formData.append("TemplateName", postArr.FileName);
            toast("Uploading","Please wait..","Uploading Template..","info","30000");
            $.ajax({
              url: "/api/plugin/ib/assessment/security/config/upload",
              type: "POST",
              data: formData,
              contentType: false,
              processData: false,
              success: function(response) {
                if (response["result"] == "Success") {
                  toast(response["result"],"",response["message"],"success","30000");
                } else if (response["result"] == "Error") {
                  toast(response["result"],"",response["message"],"danger","30000");
                } else {
                  toast("Error","","Failed to add new template","danger","30000");
                }
              },
              error: function(jqXHR, textStatus, errorThrown) {
                toast("Error","","Error uploading template","danger");
                console.error(errorThrown);
              }
            });
          }
        } else if (data["result"] == "Error") {
          toast(data["result"],"",data["message"],"danger","30000");
        } else {
          toast("Error","","Failed to add new template","danger","30000");
        }
      }).fail(function( data, status ) {
          toast("API Error","","Failed to add new template","danger","30000");
      }).always(function( data, status) {
        $("#newSATemplateModal").modal("hide");
      })
    });

    $(document).on("click", "#editSATemplateSubmit", function(event) {
      const templateFiles = $("#SAtemplatePPTX")[0].files;
      
      var id = encodeURIComponent($("#SAtemplateId").val());
      var postArr = {}
      postArr.Status = encodeURIComponent($("#SAtemplateStatus").val());
      postArr.TemplateName = encodeURIComponent($("#SAtemplateName").val());
      postArr.Description = encodeURIComponent($("#SAtemplateDescription").val());
      postArr.Orientation = encodeURIComponent($("#SAtemplateOrientation").val());
      postArr.isDefault = encodeURIComponent($("#SAtemplateSelectedByDefault")[0].checked);
      postArr.macroEnabled = encodeURIComponent($("#SAtemplateMacroEnabled")[0].checked);
      postArr.ThreatActorSlide = encodeURIComponent($("#SAtemplateThreatActorSlide").val());
      postArr.SOCInsightsSlide = encodeURIComponent($("#SAtemplateSOCInsightsSlide").val());
      if (templateFiles[0]) {
        postArr.FileName = $("#SAtemplateName").val().toLowerCase().replace(/ /g, "-");
      }
      queryAPI("PATCH", "/api/plugin/ib/assessment/security/config/"+id, postArr).done(function( data, status ) {
        if (data["result"] == "Success") {
          toast(data["result"],"",data["message"],"success");
          $("#securityAssessmentTemplateTable").bootstrapTable("refresh");
          if (templateFiles.length > 0) {
            const formData = new FormData();
            formData.append("pptx", templateFiles[0]);
            formData.append("TemplateName", postArr.FileName);
            formData.append("macroEnabled", encodeURIComponent($("#SAtemplateMacroEnabled")[0].checked));
            toast("Uploading","Please wait..","Uploading Template..","info","30000");
            $.ajax({
              url: "/api/plugin/ib/assessment/security/config/upload", // Replace with your PHP API endpoint
              type: "POST",
              data: formData,
              contentType: false,
              processData: false,
              success: function(response) {
                if (response["result"] == "Success") {
                  toast(response["result"],"",response["message"],"success","30000");
                } else if (response["result"] == "Error") {
                  toast(response["result"],"",response["message"],"danger","30000");
                } else {
                  toast("Error","","Failed to edit template","danger","30000");
                }
              },
              error: function(jqXHR, textStatus, errorThrown) {
                toast("Error","","Error uploading template","danger");
                console.error(errorThrown);
              }
            });
          }
        } else if (data["result"] == "Error") {
          toast(data["result"],"",data["message"],"danger","30000");
        } else {
          toast("Error","","Failed to edit template","danger","30000");
        }
      }).fail(function( data, status ) {
          toast("API Error","","Failed to edit template","danger","30000");
      }).always(function( data, status) {
        $("#editSATemplateModal").modal("hide");
      })
    });

    $("#threatActorConfigurationTable").bootstrapTable();
    $("#securityAssessmentTemplateTable").bootstrapTable();
  </script>
EOF;