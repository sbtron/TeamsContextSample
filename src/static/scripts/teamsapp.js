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

    var encodedWebUrl = encodeURIComponent('https://tasklist.example.com/123/456&label=Task 456');
      var encodedContext = encodeURIComponent(JSON.stringify({"subEntityId": "task456"}));
      var taskItemUrl = 'https://teams.microsoft.com/l/entity/2a4dc7a6-960d-4527-b905-a0e119db87bb/index0?webUrl=' + encodedWebUrl + '&context=' + encodedContext;
      var a= document.createElement('a');
      a.href = taskItemUrl;
      a.target = '_blank';;
      a.textContent = taskItemUrl;
      document.getElementById('deeplink').appendChild(a);

  });

  function updatev2(hubName, pageid, subpageid) {
    console.log("v2pageid: " + pageid);
    console.log("v2subpageid: " + subpageid); 
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
})();
