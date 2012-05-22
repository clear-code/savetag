window.addEventListener("load", function () {
  window.emlExporter = new EMLExporter(window);

  window.saveMessageWithTag = function () {
    var messages = gFolderDisplay.selectedMessages;
    ToggleMessageTagKey(1);
  };
}, false);
