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
              data-buttons="templateButtons"
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
                  <th data-formatter="templateActionFormatter" data-events="actionEvents">Actions</th>
                </tr>
              </thead>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- Edit Template Modal -->
  <div class="modal fade" id="editTemplateModal" tabindex="-1" role="dialog" aria-labelledby="editTemplateModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg" role="document">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="editTemplateModalLabel">Template Information</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close">
            <span aria-hidden="true"></span>
          </button>
        </div>
        <div class="modal-body" id="editTemplateModalBody">
          <div class="form-group">
            <div class="form-group" hidden>
              <input type="text" class="form-control info-field" id="templateId" aria-describedby="templateIdHelp" name="templateId" disabled>
            </div>
            <label for="templateStatus">Template Status</label>
            <select id="templateStatus" class="form-select" aria-label="Template Status" aria-describedby="templateStatusHelp">
              <option value="Active">Active</option>
              <option value="Inactive">Inactive</option>
            </select>
            <small id="templateStatusHelp" class="form-text text-muted">The current status of the template.</small>
          </div>
          <div class="form-group">
            <label for="templateName">Template Name</label>
            <input type="text" class="form-control info-field" id="templateName" aria-describedby="templateNameHelp" name="templateName">
            <small id="templateNameHelp" class="form-text text-muted">The name for the Cloud Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="templateDescription">Template Description</label>
            <input type="text" class="form-control info-field" id="templateDescription" aria-describedby="templateDescriptionHelp" name="templateDescription">
            <small id="templateDescriptionHelp" class="form-text text-muted">The description of the Cloud Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="templateOrientation" class="col-form-label">Template Orientation</label>
            <select id="templateOrientation" class="form-select" name="templateOrientation">
              <option value="Portrait" selected>Portrait</option>
              <option value="Landscape">Landscape</option>
            </select>
            <small id="templateOrientationHelp" class="form-text text-muted">The orientation of the powerpoint template.</small>
          </div>
          <div class="form-group">
            <label for="templateSelectedByDefault">Selected by Default</label>
            <div class="form-check form-switch">
              <input class="form-check-input info-field" type="checkbox" id="templateSelectedByDefault" name="templateSelectedByDefault" value="">  
            </div>
            <small id="templateSelectedByDefaultHelp" class="form-text text-muted">Enable this to select this template by default.</small>
          </div>
          <div class="form-group row">
            <label for="templatePPTX" class="col-form-label">PowerPoint Template</label>
            <div class="col-sm-5">
              <input type="file" class="form-control" id="templatePPTX" accept=".pptx" aria-describedby="templatePPTXHelp">
              <small id="templatePPTXHelp" class="form-text text-muted">Upload a PowerPoint Template.</small>
            </div>
            <div class="col-sm-5">
              <img id="imagePreviewSVG" src="" alt="PNG Image Preview" style="display:none; margin-top: 10px; max-width: 100%;">
            </div>
          </div>
          <div class="form-group">
            <label for="templateFileName">Template File Name</label>
            <input type="text" class="form-control info-field" id="templateFileName" aria-describedby="templateFileNameHelp" name="templateFileName" disabled>
            <small id="templateFileNameHelp" class="form-text text-muted">The file name for the Cloud Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="templateUploadDate">Upload Date</label>
            <input type="text" class="form-control" id="templateUploadDate" aria-describedby="templateUploadDateHelp" disabled>
            <small id="templateUploadDateHelp" class="form-text text-muted">The upload date of this template.</small>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
          <button class="btn btn-primary" id="editTemplateSubmit">Save</button>
        </div>
      </div>
    </div>
  </div>

  <!-- New Template Modal -->
  <div class="modal fade" id="newTemplateModal" tabindex="-1" role="dialog" aria-labelledby="newTemplateModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-lg" role="document">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="newTemplateModalLabel">New Template Wizard</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close">
            <span aria-hidden="true"></span>
          </button>
        </div>
        <div class="modal-body" id="newTemplateModelBody">
          <p>Enter the new template information below to add it to the list.</p>
          <div class="form-group">
            <label for="newTemplateStatus">Template Status</label>
            <select id="newTemplateStatus" class="form-select" aria-label="Template Status" aria-describedby="newTemplateStatusHelp">
              <option value="Active">Active</option>
              <option value="Inactive">Inactive</option>
            </select>
            <small id="newTemplateStatusHelp" class="form-text text-muted">The status of the new template.</small>
          </div>
          <div class="form-group">
            <label for="newTemplateName">Template Name</label>
            <input type="text" class="form-control" id="newTemplateName" aria-describedby="newTemplateNameHelp" name="newTemplateName">
            <small id="newTemplateNameHelp" class="form-text text-muted">The name for the Cloud Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="newTemplateDescription">Template Description</label>
            <input type="text" class="form-control" id="newTemplateDescription" aria-describedby="newTemplateDescriptionHelp" name="newTemplateDescription">
            <small id="newTemplateDescriptionHelp" class="form-text text-muted">The description of the Cloud Assessment Template.</small>
          </div>
          <div class="form-group">
            <label for="newTemplateOrientation" class="col-form-label">Template Orientation</label>
            <select id="newTemplateOrientation" class="form-select" name="newTemplateOrientation">
              <option value="Portrait" selected>Portrait</option>
              <option value="Landscape">Landscape</option>
            </select>
            <small id="newTemplateOrientationHelp" class="form-text text-muted">The orientation of the powerpoint template.</small>
          </div>
          <div class="form-group">
            <label for="newTemplateSelectedByDefault">Selected by Default</label>
            <div class="form-check form-switch">
              <input class="form-check-input info-field" type="checkbox" id="newTemplateSelectedByDefault" name="newTemplateSelectedByDefault" value="">  
            </div>
            <small id="newTemplateSelectedByDefaultHelp" class="form-text text-muted">Enable this to select this template by default.</small>
          </div>
          <div class="form-group row">
            <label for="newTemplatePPTX" class="col-form-label">PowerPoint Template</label>
            <div class="col-sm-5">
              <input type="file" class="form-control" id="newTemplatePPTX" accept=".pptx" aria-describedby="newTemplatePPTXHelp">
              <small id="newTemplatePPTXHelp" class="form-text text-muted">Upload a PowerPoint Template.</small>
            </div>
            <div class="col-sm-5">
              <img id="imagePreviewSVG" src="" alt="PNG Image Preview" style="display:none; margin-top: 10px; max-width: 100%;">
            </div>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
          <button class="btn btn-primary" id="newTemplateSubmit">Save</button>
        </div>
      </div>
    </div>
  </div>

  <script>
    function templateActionFormatter(value, row, index) {
      return [
        `<a class="editTemplate" title="Edit">`,
        `<i class="fa fa-pencil"></i>`,
        `</a>&nbsp;`,
        `<a class="deleteTemplate" title="Delete">`,
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
            $("#newTemplateModal").modal("show");
            $("#newTemplateModal input").val("");
          },
          attributes: {
            title: "Add a new Cloud Assessment Template",
            style: "background-color:#4bbe40;border-color:#4bbe40;"
          }
        }
      }
    }

    function listTemplate(row) {
      $("#editTemplateModal input").val("");
      $("#templateId").val(row["id"]);
      $("#templateStatus").val(row["Status"]);
      $("#templateName").val(row["TemplateName"]);
      $("#templateDescription").val(row["Description"]);
      $("#templateFileName").val(row["FileName"]);
      $("#templateOrientation").val(row["Orientation"]);
      if (String(row["isDefault"]).toLowerCase() == "true") {
        $("#templateSelectedByDefault").attr("checked",true);
      } else {
        $("#templateSelectedByDefault").attr("checked",false);
      }
      $("#templateUploadDate").val(row["Created"]);
    }

    function imagePreview(elem,img) {
      const preview = document.getElementById(elem);
      if (img) {
        preview.src = img;
        preview.style.display = "block"; // Show the image preview
      } else {
        preview.src = "";
        preview.style.display = "none"; // Hide the image preview if no file is selected
      }
    }

    window.actionEvents = {
      "click .editTemplate": function (e, value, row, index) {
        listTemplate(row);
        $("#editTemplateModal").modal("show");
      },
      "click .deleteTemplate": function (e, value, row, index) {
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

    $(document).on("click", "#newTemplateSubmit", function(event) {
      const templateFiles = $("#newTemplatePPTX")[0].files;

      var postArr = {}
      postArr.Status = encodeURIComponent($("#newTemplateStatus").val());
      postArr.TemplateName = encodeURIComponent($("#newTemplateName").val());
      postArr.Description = encodeURIComponent($("#newTemplateDescription").val());
      postArr.Orientation = encodeURIComponent($("#newTemplateOrientation").val());
      postArr.isDefault = encodeURIComponent($("#templateSelectedByDefault")[0].checked);
      if (templateFiles[0]) {
        postArr.FileName = $("#newTemplateName").val().toLowerCase().replace(/ /g, "-");
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
        $("#newTemplateModal").modal("hide");
      })
    });

    $(document).on("click", "#editTemplateSubmit", function(event) {
      const templateFiles = $("#templatePPTX")[0].files;
      
      var id = encodeURIComponent($("#templateId").val());
      var postArr = {}
      postArr.Status = encodeURIComponent($("#templateStatus").val());
      postArr.TemplateName = encodeURIComponent($("#templateName").val());
      postArr.Description = encodeURIComponent($("#templateDescription").val());
      postArr.Orientation = encodeURIComponent($("#templateOrientation").val());
      postArr.isDefault = encodeURIComponent($("#templateSelectedByDefault")[0].checked);
      if (templateFiles[0]) {
        postArr.FileName = $("#templateName").val().toLowerCase().replace(/ /g, "-");
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
        $("#editTemplateModal").modal("hide");
      })
    });

    $("#cloudAssessmentTemplateTable").bootstrapTable();
  </script>
EOF;