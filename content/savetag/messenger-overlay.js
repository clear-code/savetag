window.addEventListener("load", function () {
  var emlExporter = new EMLExporter(window);

  var prefs = Cc['@mozilla.org/preferences-service;1']
        .getService(Ci.nsIPrefBranch);

  var destination = prefs.getComplexValue("mboximport.export.server_path",
                                          Ci.nsILocalFile);
  var tagKey = prefs.getCharPref("mboximport.export.tagkey");

  window.saveMessageTag = {
    saveSelectedMessage: function () {
      try {
        ToggleMessageTag(tagKey, true);
        var selectedMessages = gFolderDisplay.selectedMessages;
        emlExporter.saveMessagesAsEMLDeferred(selectedMessages, destination)
          .next(function () {
            // alert("エクスポートが終了しました");
          }).error(function (x) {
            alert("EML Export Error: " + x);
          });
      } catch (x) {
        alert("Unknown error: " + x);
      }
    },

    displayTagList: function () {
      var tagArray = MailServices.tags.getAllTags({});
      var tagList = tagArray.map(function (tagInfo) {
        return tagInfo.tag + "  :  " + tagInfo.key;
      }).join("\n");
      alert(tagList);
    }
  };
}, false);
