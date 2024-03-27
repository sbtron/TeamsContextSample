(function () {
  "use strict";

   // Call the initialize API first
   microsoftTeams.app.initialize().then(function () {
    microsoftTeams.app.getContext().then(function (context) {
      if (context?.app?.host?.name) {
        updatev2(context.app.host.name, context.page.id, context.page.subPageId);
      }
    });
    microsoftTeams.getContext((context) => {
      updatev1(context.entityId, context.subEntityId);
    });
  });

  function updatev2(hubName, pageid, subpageid) {
    if (hubName) {
      document.getElementById("hubState").innerHTML = "in " + hubName;
      document.getElementById("v2pageid").innerHTML = "is " + pageid;
      document.getElementById("v2subpageid").innerHTML = "is " + subpageid;
    }
  }
  function updatev1(entityId, subEntityId) {
    console.log("v1entityid: " + entityId);
    document.getElementById("v1entityid").innerHTML = "is " + entityId;
    console.log("v1subentityid: " + subEntityId); 
    document.getElementById("v1subentityid").innerHTML = "is " + subEntityId;
    }

    function updateDeeplink() {
      var encodedWebUrl = encodeURIComponent('https://tasklist.example.com/123/456&label=Task 456');
      var subEntityIdvalue = document.getElementById('subEntityId').value;
      var encodedContext = encodeURIComponent(JSON.stringify({"subEntityId": subEntityIdvalue}));
          var appid = document.getElementById('appid').value;
          if (!appid) {
           console.log('No App Id. Refresh Page')
           document.getElementById('deeplink').innerHTML = 'Specify an App id and then run update deeplink';
          } 
          else {
          var taskItemUrl = 'https://teams.microsoft.com/l/entity/' + appid + 'index0?webUrl=' + encodedWebUrl + '&context=' + encodedContext;
          var a= document.createElement('a');
          a.href = taskItemUrl;
          a.target = '_blank';;
          a.textContent = taskItemUrl;
          document.getElementById('deeplink').appendChild(a);
          }
    }
    module.exports = updateDeeplink;

    function opendeeplink() {
      var encodedWebUrl = encodeURIComponent('https://tasklist.example.com/123/456&label=Task 456');
      var subEntityIdvalue = document.getElementById('subEntityId').value;
      var encodedContext = encodeURIComponent(JSON.stringify({"subEntityId": subEntityIdvalue}));
      var appid = document.getElementById('appid').value;
      if (!appid) {
        alert('Please provide an appid');
        return;
      }
      var taskItemUrl = 'https://teams.microsoft.com/l/entity/' + appid + '/index0?webUrl=' + encodedWebUrl + '&context=' + encodedContext;
      microsoftTeams.executeDeepLink(taskItemUrl);
    } 
    module.exports = opendeeplink;

})();
