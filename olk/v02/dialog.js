(function(){
    'use strict';
  
    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function(reason){
      jQuery(document).ready(function(){
  
        $('#settings-done').on('click', function() {
          const settings = {};
  
          settings.gitHubUserName = 'gitHubUserName';
  
          settings.defaultGistId = 'selectedGist';
  
            sendMessage(JSON.stringify(settings));
          }
        });
      });
    };
  

  

  
    function sendMessage(message) {
      Office.context.ui.messageParent(message);
    }
  
  })();
