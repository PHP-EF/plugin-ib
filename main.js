var maxDaysApart = 31;
var today = new Date();
var maxPastDate = new Date(today);
maxPastDate.setDate(today.getDate() - 31);

flatpickr("#reportingStartAndEndDate", {
  mode: "range",
  maxDate: today,
  enableTime: true,
  dateFormat: "Y-m-d H:i"
});

flatpickr("#anonReportingStartAndEndDate", {
  mode: "range",
  maxDate: today,
  enableTime: true,
  dateFormat: "Y-m-d H:i"
});

function apiKeyBtn(elem) {
  if ($(elem).hasClass('fa-save')) {
    saveAPIKey($(elem).prev().val());
  } else if ($(elem).hasClass('fa-trash')) {
    removeAPIKey();
  }
}

function saveAPIKey(key) {
    queryAPI("POST", "/api/auth/crypt", {key: key}).done(function( data, status ) {
        if (data.result == 'Success') {
            setCookie('crypt',data.data,7);
            checkAPIKey();
            toast("Success","","Saved API Key.","success","30000");
        } else {
            toast(data.result,"","Unable to save API Key.","danger","30000");
        }
    }).fail(function( data, status ) {
        toast("API Error","","Unable to save API Key.","danger","30000");
    })
}

function checkAPIKey() {
    if (getCookie('crypt')) {
        $('.APIKey').prop('disabled',true).attr('placeholder','== Using Saved API Key ==').val('');
        $(".saveBtn").removeClass('fa-save').addClass('fa-trash')
        checkInput('saved');
    } else {
        $('.APIKey').prop('disabled',false).attr('placeholder','Enter API Key');
        $(".saveBtn").removeClass('fa-trash').addClass('fa-save')
    }
}

function removeAPIKey() {
    setCookie('crypt',null,-1);
    checkAPIKey();
    toast("Success","","Removed API Key.","success","30000");
}

function checkInput(text) {
    if (text) {
        $(".saveBtn").addClass("saveBtnShow");
    } else {
        $(".saveBtn").removeClass("saveBtnShow");
    }
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

// ** Populate the API Key input field ** //
// Function to be called when specific elements are created
function runOnAPIKeyCreation() {
  checkAPIKey();
  // Add your custom functions here
}

// ** DNS Toolbox ** //
function requestSource(source){
  switch (source) {
      case 'google':
          return "Google DNS";
      case 'cloudflare':
          return "Cloudflare DNS";
  }
}

function showLoadingDNS() {
  document.querySelector('.loading-icon').style.display = 'block';
  document.querySelector('.loading-div').style.display = 'block';
}
  
function hideLoadingDNS() {
  document.querySelector('.loading-icon').style.display = 'none';
  document.querySelector('.loading-div').style.display = 'none';
}

function returnDnsDetails(domain, callType, port, source, doh) {
  $('#txtHint, .info').html('');

  console.log(domain, callType, port, source, doh);

  var queryParams = 'domain=' + encodeURIComponent(domain) + '&request=' + callType + '&source=' + source;
  if (port) {
    queryParams += '&port=' + encodeURIComponent(port);
  }
  if (doh) {
    queryParams += '&dohServer=' + encodeURIComponent(doh);
  }

  queryAPI("GET", "/api/dnstoolbox?" + queryParams).done(function( data ) {
      if (data['result'] == 'Error') {
          toast(data['result'],"",data['Message'],"danger","30000");
          hideLoadingDNS();
      } else {
          const columns = [];
          switch(callType) {
              case 'port':
                  columns.push({
                      field: 'hostname',
                      title: 'Hostname',
                      sortable: true
                  },{
                      field: 'port',
                      title: 'Port',
                      sortable: true
                  },
                  {
                      field: 'result',
                      title: 'Status',
                      sortable: true
                  });
                  break;
              case 'ptr':
                  columns.push({
                      field: 'ip',
                      title: 'IP Address',
                      sortable: true
                  },{
                      field: 'hostname',
                      title: 'Hostname',
                      sortable: true
                  });
                  break;
              default:
                  columns.push({
                      field: 'hostname',
                      title: 'Hostname',
                      sortable: true
                  },{
                      field: 'type',
                      title: 'Type',
                      sortable: true
                  },
                  {
                      field: 'TTL',
                      title: 'TTL',
                      sortable: true
                  },
                  {
                      field: 'class',
                      title: 'Class',
                      sortable: true
                  });
                  break;
          }


          if ((["a","aaaa","all"]).includes(callType)) {
              columns.push({
                  field: 'IPAddress',
                  title: 'IP Address',
                  sortable: true
              })
          }
          if ((["mx","txt","dmarc","all","nameserver","soa","cname"]).includes(callType)) {
              columns.push({
                  field: 'data',
                  title: 'Data',
                  sortable: true
              })
          }

          $('#dnsResponseTable').bootstrapTable('destroy');
          $('#dnsResponseTable').bootstrapTable({
              data: data,
              dataField: 'data',
              sortable: true,
              pagination: true,
              search: true,
              showExport: true,
              exportTypes: ['json', 'xml', 'csv', 'txt', 'excel', 'sql'],
              showColumns: true,
              columns: columns
          });
          hideLoadingDNS();
          return false;
      }
  }).fail(function() {
    toast("API Error","","Unknown API Error","danger","30000");
  }).always(function() {
    hideLoadingDNS();
  });
}

function showAdditionalFields() {
  var file = $("#file");
  var source = $("#source");;
  var port = $("#port-container");
  var doh = $("#doh-container");

  source.attr('disabled', false);
  source.parent().css('display','block');
  port.parent().css('display','none');
  doh.parent().css('display','none');
  $('#port, #doh').val('');

  switch (file.val()) {
      case 'port':
          port.parent().css('display','block');
          source.attr('disabled', true);
          source.parent().css('display','none');
          break;
      case 'reverseLookup':
          break;
      default:
          switch (source.val()) {
            case 'doh':
              doh.parent().css('display','block');
                break;
            default:
                break;
          }
          break;
  }
}

$('#copyLink').on('click',function(elem) {
  elem.preventDefault();
  var domain = $('#domain').val();
  var file = $('#file').val();
  var source = $('#source').val();

  var text = window.location.origin+"/pages/tools/dnstoolbox.php?domain="+domain+"&type="+file+"&location="+source;

  // Copy the text inside the text field
  navigator.clipboard.writeText(text);

  // Alert the copied text
  toast("Info","","Copied link to clipboard","primary");
});
// ** DNS Toolbox ** //

// MutationObserver callback function
function mutationCallback(mutationsList, observer) {
  for (let mutation of mutationsList) {
      if (mutation.type === 'childList') {
          for (let node of mutation.addedNodes) {
              if (node.nodeType === Node.ELEMENT_NODE) {
                  // Check if the created element matches your criteria
                  if (node.matches('.APIKey')) {
                    runOnAPIKeyCreation();
                  }
                  // Check for nested elements
                  const nestedElements = node.querySelectorAll('.APIKey');
                  nestedElements.forEach(runOnAPIKeyCreation);
              }
          }
      }
  }
}

runOnAPIKeyCreation();

// Create a MutationObserver instance
const observer = new MutationObserver(mutationCallback);

// Start observing the target node for configured mutations
const targetNode = document.body; // You can change this to a more specific target
const mutationObserverConfig = { childList: true, subtree: true };
observer.observe(targetNode, mutationObserverConfig);
