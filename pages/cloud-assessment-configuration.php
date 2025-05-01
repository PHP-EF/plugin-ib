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
        <h2 class="h3 mb-4 page-title">Cloud Assessment Report Generator Configuration</h2>
          <div class="card border-secondary">
            <div class="card-title">
              <h5>Template Configuration</h5>
              <p>Use the following to configure the template for the Cloud Assessment Report Generator.</p>
            </div>
            <table  data-url="/api/plugin/ib/assessment/cloud/config"
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
              data-buttons="CAtemplateButtons"
              data-buttons-order="btnAddTemplate,refresh"
              class="table table-striped" id="cloudAssessmentTemplateTable">

              <thead>
                <tr>
                  <th data-field="state" data-checkbox="true"></th>
                  <th data-field="Status" data-sortable="true">Status</th>
                  <th data-field="TemplateName" data-sortable="true">Name</th>
                  <th data-field="Description" data-sortable="true">Description</th>
                  <th data-field="FileName" data-sortable="true">File Name</th>
                  <th data-field="Created" data-sortable="true">Upload Date</th>
                  <th data-formatter="CAtemplateActionFormatter" data-events="actionEvents">Actions</th>
                </tr>
              </thead>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- Edit Template Modal -->
  <div class="modal fade" id="editCATemplateModal" tabindex="-1" role="dialog" aria-labelledby="editCATemplateModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg" role="document">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="editCATemplateModalLabel">Template Information</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close">
            <span aria-hidden="true"></span>
          </button>
        </div>
        <div class="modal-body" id="editCATemplateModalBody">
          <div class="form-group">
            <div class="form-group" hidden>
              <input type="text" class="form-control info-field" id="CAtemplateId" aria-describedby="CAtemplateIdHelp" name="CAtemplateId" disabled>
            </div>
            <label for="CAtemplateStatus">Template Status</label>
            <select id="CAtemplateStatus" class="form-select" aria-label="Template Status" aria-describedby="CAtemplateStatusHelp">
              <option value="Active">Active</option>
              <option value="Inactive">Inactive</option>
            </select>
            <small id="CAtemplateStatusHelp" class="form-text text-muted">The current status of the template.</small>
          </div>
          <div class="form-group">
            <label for="CAtemplateName">Template Name</label>
            <input type="text" class="form-control info-field" id="CAtemplateName" aria-describedby="CAtemplateNameHelp" name="CAtemplateName">
            <small id="CAtemplateNameHelp" class="form-text text-muted">The name for the Cloud Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="CAtemplateDescription">Template Description</label>
            <input type="text" class="form-control info-field" id="CAtemplateDescription" aria-describedby="CAtemplateDescriptionHelp" name="CAtemplateDescription">
            <small id="CAtemplateDescriptionHelp" class="form-text text-muted">The description of the Cloud Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="CAtemplateOrientation" class="col-form-label">Template Orientation</label>
            <select id="CAtemplateOrientation" class="form-select" name="CAtemplateOrientation">
              <option value="Portrait" selected>Portrait</option>
              <option value="Landscape">Landscape</option>
            </select>
            <small id="CAtemplateOrientationHelp" class="form-text text-muted">The orientation of the powerpoint template.</small>
          </div>
          <div class="form-group">
            <label for="CAtemplateSelectedByDefault">Selected by Default</label>
            <div class="form-check form-switch">
              <input class="form-check-input info-field" type="checkbox" id="CAtemplateSelectedByDefault" name="CAtemplateSelectedByDefault" value="">  
            </div>
            <small id="CAtemplateSelectedByDefaultHelp" class="form-text text-muted">Enable this to select this template by default.</small>
          </div>
          <div class="form-group row">
            <label for="CAtemplatePPTX" class="col-form-label">PowerPoint Template</label>
            <div class="col-sm-5">
              <input type="file" class="form-control" id="CAtemplatePPTX" accept=".pptx" aria-describedby="CAtemplatePPTXHelp">
              <small id="CAtemplatePPTXHelp" class="form-text text-muted">Upload a PowerPoint Template.</small>
            </div>
            <div class="col-sm-5">
              <img id="imagePreviewSVG" src="" alt="PNG Image Preview" style="display:none; margin-top: 10px; max-width: 100%;">
            </div>
          </div>
          <div class="form-group">
            <label for="CAtemplateFileName">Template File Name</label>
            <input type="text" class="form-control info-field" id="CAtemplateFileName" aria-describedby="CAtemplateFileNameHelp" name="CAtemplateFileName" disabled>
            <small id="CAtemplateFileNameHelp" class="form-text text-muted">The file name for the Cloud Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="CAtemplateUploadDate">Upload Date</label>
            <input type="text" class="form-control" id="CAtemplateUploadDate" aria-describedby="CAtemplateUploadDateHelp" disabled>
            <small id="CAtemplateUploadDateHelp" class="form-text text-muted">The upload date of this template.</small>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
          <button class="btn btn-primary" id="editCATemplateSubmit">Save</button>
        </div>
      </div>
    </div>
  </div>

  <!-- New Template Modal -->
  <div class="modal fade" id="newCATemplateModal" tabindex="-1" role="dialog" aria-labelledby="newCATemplateModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg" role="document">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="newCATemplateModalLabel">New Template Wizard</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close">
            <span aria-hidden="true"></span>
          </button>
        </div>
        <div class="modal-body" id="newCATemplateModelBody">
          <p>Enter the new template information below to add it to the list.</p>
          <div class="form-group">
            <label for="newCATemplateStatus">Template Status</label>
            <select id="newCATemplateStatus" class="form-select" aria-label="Template Status" aria-describedby="newCATemplateStatusHelp">
              <option value="Active">Active</option>
              <option value="Inactive">Inactive</option>
            </select>
            <small id="newCATemplateStatusHelp" class="form-text text-muted">The status of the new template.</small>
          </div>
          <div class="form-group">
            <label for="newCATemplateName">Template Name</label>
            <input type="text" class="form-control" id="newCATemplateName" aria-describedby="newCATemplateNameHelp" name="newCATemplateName">
            <small id="newCATemplateNameHelp" class="form-text text-muted">The name for the Cloud Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="newCATemplateDescription">Template Description</label>
            <input type="text" class="form-control" id="newCATemplateDescription" aria-describedby="newCATemplateDescriptionHelp" name="newCATemplateDescription">
            <small id="newCATemplateDescriptionHelp" class="form-text text-muted">The description of the Cloud Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="newCATemplateOrientation" class="col-form-label">Template Orientation</label>
            <select id="newCATemplateOrientation" class="form-select" name="newCATemplateOrientation">
              <option value="Portrait" selected>Portrait</option>
              <option value="Landscape">Landscape</option>
            </select>
            <small id="newCATemplateOrientationHelp" class="form-text text-muted">The orientation of the powerpoint template.</small>
          </div>
          <div class="form-group">
            <label for="newCATemplateSelectedByDefault">Selected by Default</label>
            <div class="form-check form-switch">
              <input class="form-check-input info-field" type="checkbox" id="newCATemplateSelectedByDefault" name="newCATemplateSelectedByDefault" value="">  
            </div>
            <small id="newCATemplateSelectedByDefaultHelp" class="form-text text-muted">Enable this to select this template by default.</small>
          </div>
          <div class="form-group row">
            <label for="newCATemplatePPTX" class="col-form-label">PowerPoint Template</label>
            <div class="col-sm-5">
              <input type="file" class="form-control" id="newCATemplatePPTX" accept=".pptx" aria-describedby="newCATemplatePPTXHelp">
              <small id="newCATemplatePPTXHelp" class="form-text text-muted">Upload a PowerPoint Template.</small>
            </div>
            <div class="col-sm-5">
              <img id="imagePreviewSVG" src="" alt="PNG Image Preview" style="display:none; margin-top: 10px; max-width: 100%;">
            </div>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
          <button class="btn btn-primary" id="newCATemplateSubmit">Save</button>
        </div>
      </div>
    </div>
  </div>

  <script>
    function CAtemplateActionFormatter(value, row, index) {
      return [
        `<a class="editCATemplate" title="Edit">`,
        `<i class="fa fa-pencil"></i>`,
        `</a>&nbsp;`,
        `<a class="deleteCATemplate" title="Delete">`,
        `<i class="fa fa-trash"></i>`,
        `</a>`
      ].join("")
    }

    function templateButtons() {
      return {
        btnAddTemplate: {
          text: "Add Cloud Assessment Template",
          icon: "bi-plus-lg",
          event: function() {
            $("#newCATemplateModal").modal("show");
            $("#newCATemplateModal input").val("");
          },
          attributes: {
            title: "Add a new Cloud Assessment Template",
            style: "background-color:#4bbe40;border-color:#4bbe40;"
          }
        }
      }
    }

    function listCATemplate(row) {
      $("#editCATemplateModal input").val("");
      $("#CAtemplateId").val(row["id"]);
      $("#CAtemplateStatus").val(row["Status"]);
      $("#CAtemplateName").val(row["TemplateName"]);
      $("#CAtemplateDescription").val(row["Description"]);
      $("#CAtemplateFileName").val(row["FileName"]);
      $("#CAtemplateOrientation").val(row["Orientation"]);
      if (String(row["isDefault"]).toLowerCase() == "true") {
        $("#CAtemplateSelectedByDefault").attr("checked",true);
      } else {
        $("#CAtemplateSelectedByDefault").attr("checked",false);
      }
      $("#CAtemplateUploadDate").val(row["Created"]);
    }

    window.actionEvents = {
      "click .editCATemplate": function (e, value, row, index) {
        listCATemplate(row);
        $("#editCATemplateModal").modal("show");
      },
      "click .deleteCATemplate": function (e, value, row, index) {
        if(confirm("Are you sure you want to delete "+row.TemplateName+" from the list of Templates? This is irriversible.") == true) {
          queryAPI("DELETE", "/api/plugin/ib/assessment/cloud/config/"+row.id).done(function( data, status ) {
            if (data["result"] == "Success") {
              toast(data["result"],"",data["message"],"success");
              $("#cloudAssessmentTemplateTable").bootstrapTable("refresh");
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

    $(document).on("click", "#newCATemplateSubmit", function(event) {
      const templateFiles = $("#newCATemplatePPTX")[0].files;

      var postArr = {}
      postArr.Status = encodeURIComponent($("#newCATemplateStatus").val());
      postArr.TemplateName = encodeURIComponent($("#newCATemplateName").val());
      postArr.Description = encodeURIComponent($("#newCATemplateDescription").val());
      postArr.Orientation = encodeURIComponent($("#newCATemplateOrientation").val());
      postArr.isDefault = encodeURIComponent($("#CAtemplateSelectedByDefault")[0].checked);
      if (templateFiles[0]) {
        postArr.FileName = $("#newCATemplateName").val().toLowerCase().replace(/ /g, "-");
      }
      queryAPI("POST", "/api/plugin/ib/assessment/cloud/config", postArr).done(function( data, status ) {
        if (data["result"] == "Success") {
          toast(data["result"],"",data["message"],"success");
          $("#cloudAssessmentTemplateTable").bootstrapTable("refresh");
          if (templateFiles.length > 0) {
            const formData = new FormData();
            formData.append("pptx", templateFiles[0]);
            formData.append("TemplateName", postArr.FileName);
            toast("Uploading","Please wait..","Uploading Template..","info","30000");
            $.ajax({
              url: "/api/plugin/ib/assessment/cloud/config/upload",
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
        $("#newCATemplateModal").modal("hide");
      })
    });

    $(document).on("click", "#editCATemplateSubmit", function(event) {
      const templateFiles = $("#CAtemplatePPTX")[0].files;
      
      var id = encodeURIComponent($("#CAtemplateId").val());
      var postArr = {}
      postArr.Status = encodeURIComponent($("#CAtemplateStatus").val());
      postArr.TemplateName = encodeURIComponent($("#CAtemplateName").val());
      postArr.Description = encodeURIComponent($("#CAtemplateDescription").val());
      postArr.Orientation = encodeURIComponent($("#CAtemplateOrientation").val());
      postArr.isDefault = encodeURIComponent($("#CAtemplateSelectedByDefault")[0].checked);
      if (templateFiles[0]) {
        postArr.FileName = $("#CAtemplateName").val().toLowerCase().replace(/ /g, "-");
      }
      queryAPI("PATCH", "/api/plugin/ib/assessment/cloud/config/"+id, postArr).done(function( data, status ) {
        if (data["result"] == "Success") {
          toast(data["result"],"",data["message"],"success");
          $("#cloudAssessmentTemplateTable").bootstrapTable("refresh");
          if (templateFiles.length > 0) {
            const formData = new FormData();
            formData.append("pptx", templateFiles[0]);
            formData.append("TemplateName", postArr.FileName);
            toast("Uploading","Please wait..","Uploading Template..","info","30000");
            $.ajax({
              url: "/api/plugin/ib/assessment/cloud/config/upload", // Replace with your PHP API endpoint
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
        $("#editCATemplateModal").modal("hide");
      })
    });

    $("#cloudAssessmentTemplateTable").bootstrapTable();
  </script>
EOF;