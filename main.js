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



// ** Populate the API Key input field ** //
// Function to be called when specific elements are created
function runOnAPIKeyCreation() {
  checkAPIKey();
  $('.saveBtn').click(function(){
    if ($('.saveBtn').hasClass('fa-save')) {
      saveAPIKey($('.APIKey').val());
    } else if ($('.saveBtn').hasClass('fa-trash')) {
      removeAPIKey();
    }
  });
  // Add your custom functions here
}

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
const config = { childList: true, subtree: true };
observer.observe(targetNode, config);
