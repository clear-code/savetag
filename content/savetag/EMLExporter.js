var { classes: Cc, interfaces: Ci, utils: Cu, results: Cr } = Components;

var NS_MSG_ERROR_FOLDER_SUMMARY_OUT_OF_DATE = 0x80550005;
var NS_MSG_ERROR_FOLDER_SUMMARY_MISSING = 0x80550006;

Util.DEBUG = true;

Deferred.onerror = function (x) {
  dump(x.toString() + "\n");
  // alert(x.toString());
  if (x.stack) {
    dump(x.stack.toString() + "\n");
    // alert(x.stack);
  }
};

function ProgressManager(onCountChange) {
  this.onCountChange = onCountChange;
  this.currentTaskNumber = this.totalTaskCount = 0;
}

ProgressManager.prototype = {
  isFinished: function () {
    return this.currentTaskNumber >= this.totalTaskCount;
  },

  countUp: function () {
    // this.currentTaskNumber++;
    // this.onCountChange();
  },

  resetTaskCount: function (totalTaskCount) {
    this.totalTaskCount = totalTaskCount;
    this.currentTaskNumber = 0;
    this.onCountChange();
  },

  getProgressText: function () {
    if (this.totalTaskCount > 0)
      return "(" + (this.currentTaskNumber + 1) + " / " + this.totalTaskCount + ")";
    else
      return "";
  }
};

function EMLExporter(global, elements) {
  elements = elements || {};

  this.global = global;
  this.strbundle = Cc["@mozilla.org/intl/stringbundle;1"]
                    .getService(Ci.nsIStringBundleService)
                    .createBundle("chrome://savetag/locale/savetag.properties");
  this.cancelButton = elements.cancelButton;
}

EMLExporter.isThunderbird2 = !Ci.nsMsgFolderFlags;

EMLExporter.prototype = {
  set displayCancelButton(display) {
    if (this.cancelButton)
      this.cancelButton.hidden = !display;
  },
  get document() { return this.global.document; },
  get progressElement() { return this.document.getElementById("eml-progress-reporter"); },
  get selectedMessageFolders() {
    return this.global.GetSelectedMsgFolders();
  },
  get currentMessageFolder() {
    var selectedMessageFolders = this.selectedMessageFolders;
    return selectedMessageFolders ? selectedMessageFolders[0] : null;
  },
  get currentMessageFolderIsServer() {
    return this.currentMessageFolder && this.currentMessageFolder.isServer;
  },

  get accounts() {
    let acctMgr = Cc["@mozilla.org/messenger/account-manager;1"]
          .getService(Ci.nsIMsgAccountManager);
    let accounts = Util.toArray(acctMgr.accounts, Ci.nsIMsgAccount);
    // Bug 41133 workaround
    accounts = accounts.filter(function fix(account) { return account.incomingServer; });

    // Don't show deferred pop accounts
    accounts = accounts.filter(function isNotDeferred(account) {
      let server = account.incomingServer;
      return !(server instanceof Ci.nsIPop3IncomingServer && server.deferredToAccount);
    });

    return accounts;
  },

  set displayProgressElement(display) {
    this.progressElement.hidden = !display;
  },

  setNotificationMessage: function (message) {
    this.progressElement.label = message;
  },

  updateNotification: function () {
    this.setNotificationMessage(this.getCurrentNotificationText());
  },

  exportingMessages: function () {
    return !this.folderExportingProgress.isFinished();
  },

  getCurrentNotificationText: function () {
    return this.strbundle.formatStringFromName("exportingProgressMessage", [
      this.folderExportingProgress.getProgressText(),
      this.messageExportingProgress.getProgressText()
    ], 2);
  },

  onTaskCountChange: function () {
    // this.updateNotification();
  },

  get folderExportingProgress() {
    if (!this._folderExportingProgress) {
      var self = this;
      this._folderExportingProgress = new ProgressManager(function () {
        self.onTaskCountChange();
      });
    }
    return this._folderExportingProgress;
  },

  get messageExportingProgress() {
    if (!this._messageExportingProgress) {
      var self = this;
      this._messageExportingProgress = new ProgressManager(function () {
        self.onTaskCountChange();
      });
    }
    return this._messageExportingProgress;
  },

  promptDestination: function () {
    var destination = null;

    if (!destination) {
      let fp = Cc["@mozilla.org/filepicker;1"].createInstance(Ci.nsIFilePicker);
      fp.init(this.global, this.strbundle.GetStringFromName("selectDestination"), Ci.nsIFilePicker.modeGetFolder);

      if (fp.show() === Ci.nsIFilePicker.returnOK)
        destination = fp.file;
      else
        destination = null;
    }

    return destination;
  },

  nsMsgFolderFlags: EMLExporter.isThunderbird2 ? {
    // Prior to Thunderbird 3.0, we don't have nsMsgFolderFlags
    Inbox     : 0x00001000,
    Drafts    : 0x00000400,
    Trash     : 0x00000100,
    SentMail  : 0x00000200,
    Templates : 0x00400000,
    Junk      : 0x40000000,
    Archive   : 0x00004000
  } : {
    Inbox     : Ci.nsMsgFolderFlags.Inbox,
    Drafts    : Ci.nsMsgFolderFlags.Drafts,
    Trash     : Ci.nsMsgFolderFlags.Trash,
    SentMail  : Ci.nsMsgFolderFlags.SentMail,
    Templates : Ci.nsMsgFolderFlags.Templates,
    Junk      : Ci.nsMsgFolderFlags.Junk,
    Archive   : Ci.nsMsgFolderFlags.Archive
  },

  isSmartFolder: function (messageFolder) {
    return messageFolder.flags &
      (this.nsMsgFolderFlags.Inbox     |
       this.nsMsgFolderFlags.Drafts    |
       this.nsMsgFolderFlags.Trash     |
       this.nsMsgFolderFlags.SentMail  |
       this.nsMsgFolderFlags.Templates |
       this.nsMsgFolderFlags.Junk      |
       this.nsMsgFolderFlags.Archive);
  },

  isMessageFolderVirtual: function (messageFolder) {
    return !!(messageFolder && messageFolder.flags & 0x00000020 /* Ci.nsMsgFolderFlags.Virtual */);
  },

  getFTVItemByMessageFolder: function (messageFolder) {
    if (typeof gFolderTreeView === "undefined")
      return null;

    for (let [, ftvItem] in Iterator(gFolderTreeView._rowMap)) {
      if (ftvItem._folder === messageFolder)
        return ftvItem;
    }

    return null;
  },

  getMessageFolderName: function (messageFolder) {
    if (this.isSmartFolder(messageFolder))
      return this.getSmartFolderName(messageFolder);
    else
      return messageFolder.name;
  },

  getSmartFolderName: function (smartFolder) {
    let ftvItem = this.getFTVItemByMessageFolder(smartFolder);
    if (!ftvItem)
      return smartFolder.name;

    let smartFolderName;
    if (ftvItem.useServerNameOnly) {
      smartFolderName = ftvItem._folder.server.prettyName;
    } else {
      smartFolderName = ftvItem._folder.abbreviatedName;
      if (ftvItem.addServerName)
        smartFolderName += " - " + ftvItem._folder.server.prettyName;
    }

    return smartFolderName;
  },

  messageFolderToLocalFile: function (messageFolder) {
    return Util.getFile(messageFolder.filePath);
  },

  messageURIToMessageHdr: function (uri) {
    return this.global.messenger.messageServiceFromURI(uri).messageURIToMsgHdr(uri);
  },

  /**
   * @throws nsIException NS_MSG_ERROR_FOLDER_SUMMARY_OUT_OF_DATE
   * @throws nsIException NS_MSG_ERROR_FOLDER_SUMMARY_MISSING
   */
  getMessageArrayFromFolder: function (messageFolder, startTime) {
    startTime = startTime || Date.now();
    var self = this;
    return Deferred.next(function() {
      var messages = null;
      if (messageFolder.getMessages) { // Thunderbird 2.*
        messages = messageFolder.getMessages(null);
      } else { // Thunderbird 3.0 and later
        messages = messageFolder.messages;
      }
      return messages ? Util.toArray(messages, Ci.nsIMsgDBHdr) : [];
    }).error(function(error) {
      if (error && error instanceof Ci.nsIException) {
        let messageFolderLocation = messageFolder.filePath;
        let additionalInformation = [messageFolder.prettiestName, messageFolderLocation.path];
        let now = Date.now();
        switch (error.result) {
          case NS_MSG_ERROR_FOLDER_SUMMARY_OUT_OF_DATE:
            if (now - startTime < self.maxReparseTime) {
              return self.reparseFolder(messageFolder)
                         .next(function() {
                           return self.getMessageArrayFromFolder(messageFolder, startTime);
                         });
            }
            error = self.strbundle.formatStringFromName("updatingFolderSummaries", additionalInformation, additionalInformation.length);
            break;
          case NS_MSG_ERROR_FOLDER_SUMMARY_MISSING:
            if (now - startTime < self.maxReparseTime) {
              return self.reparseFolder(messageFolder)
                         .next(function() {
                           return self.getMessageArrayFromFolder(messageFolder, startTime);
                         });
            }
            error = self.strbundle.formatStringFromName("missingFolderSummaries", additionalInformation, additionalInformation.length);
            break;
        }
      }
      throw error;
    });
  },
  get maxReparseTime() {
    return Cc['@mozilla.org/preferences;1']
             .getService(Ci.nsIPrefBranch)
             .getIntPref("extensions.thunderbirdexporttool@mitsubishielectric.co.jp.maxReparseTime");
  },
  reparseFolder: function (folder) {
    var deferred = new Deferred();
    var self = this;
    Deferred.next(function() {
      if (!folder)
        return deferred.fail(new Error("missing folder"));

      try {
        folder = folder.QueryInterface(Ci.nsIMsgLocalMailFolder);
      } catch (x) {
        return deferred.fail(new Error("not a local folder"));
      }

      try {
        folder.getDatabaseWithReparse(null, self.global.msgWindow);
      } catch(x) {
        Util.log(x);
        // already parsing!
      }
      Deferred.wait(0.5).next(function() {
        deferred.call();
      });
    }).error(function(e) {
      Util.log(e);
      deferred.call();
    });
    return deferred;
  },

  getAllFoldersWithFlags: function (flags) {
    let allFoldersWithFlags = [];

    for (let [, account] in Iterator(this.accounts)) {
      let foldersWithFlags = account.incomingServer.rootFolder.getFoldersWithFlags(flags);
      for (let [, folderWithFlags] in Iterator(Util.toArray(foldersWithFlags.enumerate(), Ci.nsIMsgFolder))) {
        allFoldersWithFlags.push(folderWithFlags);

        // Add sub-folders of Sent and Archive to the result.
        // if (deep && (aFolderFlag & (nsMsgFolderFlags.SentMail | nsMsgFolderFlags.Archive)))
        //   this.addSubFolders(folderWithFlag, folders);
      }
    }

    return allFoldersWithFlags;
  },

  getAllSubFoldersForSmartFolder: function (smartFolder) {
    let smartFtvItem = this.getFTVItemByMessageFolder(smartFolder);
    let subFolders = [];

    if (smartFtvItem && smartFtvItem.children) {
      for (let [, childFtvItem] in Iterator(smartFtvItem.children)) {
        subFolders.push(childFtvItem._folder);
      }
    }

    return subFolders;
  },

  getAllSubFolders: function (messageFolder) {
    if (messageFolder.GetSubFolders) {
      // <= Thunderbird 2.0
      let subFolderEnumerator = messageFolder.GetSubFolders();
      let subFolders = [];

      while (true) {
        try {
          let nextItem  = subFolderEnumerator.currentItem();
          let subFolder = nextItem.QueryInterface(Ci.nsIMsgFolder);

          subFolders.push(subFolder);
          subFolderEnumerator.next();
        } catch (x) {
          break;
        }
      }

      return subFolders;
    } else if (this.isSmartFolder(messageFolder)) {
      return this.getAllSubFoldersForSmartFolder(messageFolder);
    } else {
      return Util.toArray(messageFolder.subFolders, Ci.nsIMsgFolder);
    }
  },

  // calculates a count of all descendent folders (includes self)
  getAllDescendentFoldersCount: function (messageFolder) {
    let subFoldersCount = 1;    // self
    for (let [, subFolder] in Iterator(this.getAllSubFolders(messageFolder))) {
      subFoldersCount += this.getAllDescendentFoldersCount(subFolder);
    }
    return subFoldersCount;
  },

  // See http://mxr.mozilla.org/comm-central/source/mailnews/base/public/nsMsgFolderFlags.idl
  extractMailFolderFlags: function (flags) {
    var extractedFlags = [];

    var masks = [
      0x00000F00,
      0x0000F000,
      0x000F0000,
      0x00F00000,
      0x0F000000,
      0xF0000000
    ];

    for (let [, mask] in Iterator(masks)) {
      if (flags & mask)
        extractedFlags.push(flags & mask);
    }

    return extractedFlags;
  },

  exportEMLInternalDeferred: function (recursive) {
    var self = this;

    var destination = this.promptDestination();
    if (!destination) {
      // canceled
      return Deferred.next(function () {});
    }

    if (!destination.exists() || !destination.isWritable()) {
      return Deferred.next(function () {
        let errorMessageKeyName = !destination.exists()
              ? "destinationNotExist"
              : "destinationNotWritable";
        let errorMessage = self.strbundle.GetStringFromName(errorMessageKeyName);
        Util.alert(self.strbundle.GetStringFromName("errorMessageTitle"),
                   errorMessage);
      });
    }

    this.onStartExport(recursive);

    var createdDirectories = [];
    var errorCaughtInExportation = null;

    function cleanCreatedDirectories() {
      createdDirectories.forEach(function (createdDirectory) {
        try {
          createdDirectory.remove(true);
        } catch (x) {
          Util.log("Failed to remove a folder => " + x);
        }
      });
    }

    var selectedMessageFolders = this.selectedMessageFolders;

    var exportCanceled = false;

    var currentExportProcess;
    var exportEMLDeferred = Deferred.loop(selectedMessageFolders.length, function (index) {
      // If some error has occured, skip rest directories
      if (errorCaughtInExportation || exportCanceled)
        return false;

      var targetMessageFolder = selectedMessageFolders[index];
      Util.log("targetMessageFolder => " + targetMessageFolder);
      currentExportProcess = self.exportAllEMLInLocalFolderDeferred(
        destination, targetMessageFolder, recursive
      );

      return currentExportProcess.next(function (exportInformation) {
        if (exportInformation) {
          createdDirectories.push(exportInformation.createdDirectory);
          if (exportInformation.error) {
            errorCaughtInExportation = exportInformation.error; // Save what error occured
          }
        }
      });
    }).error(function (unhandledError) {
      // Catch unhandled error
      errorCaughtInExportation = unhandledError;
    }).next(function () {
      if (exportCanceled)
        return;

      // finish exporting
      self.onFinishExport();

      if (errorCaughtInExportation) {
        // Export failed. Remove created directories.
        Util.alert(self.strbundle.GetStringFromName("errorMessageTitle"), errorCaughtInExportation + "");
        cleanCreatedDirectories();
      }
    });

    exportEMLDeferred.canceller = function () {
      exportCanceled = true;
      currentExportProcess.cancel();
      self.onFinishExport();
    };

    return exportEMLDeferred;
  },

  countExportFolders: function (messageFolders, recursive) {
    var exportFolderCount;

    if (recursive) {
      exportFolderCount = 0;
      // add all subfolder
      for (let [, messageFolder] in Iterator(messageFolders)) {
        exportFolderCount += this.getAllDescendentFoldersCount(messageFolder);
      }
    } else {
      exportFolderCount = messageFolders.length;
    }

    return exportFolderCount;
  },

  onStartExport: function (recursive) {
    this.displayCancelButton = true;
    this.displayProgressElement = true;

    var exportFolderCount = this.countExportFolders(this.selectedMessageFolders, recursive);
    this.folderExportingProgress.resetTaskCount(exportFolderCount);
  },

  onFinishExport: function () {
    this.displayCancelButton = false;
    this.displayProgressElement = false;

    this.folderExportingProgress.resetTaskCount(0);
    this.messageExportingProgress.resetTaskCount(0);
  },

  askCancelExport: function () {
    if (!this.exportingMessages)
      return;

    var cancelExport = 0 == Util.confirmEx(
      window,
      this.strbundle.GetStringFromName("askCancelExportTitle"),
      this.strbundle.GetStringFromName("askCancelExport"),
      Util.prompts.BUTTON_POS_0 * Util.prompts.BUTTON_TITLE_YES +
        Util.prompts.BUTTON_POS_1 * Util.prompts.BUTTON_TITLE_NO
    );
    if (cancelExport)
      this.cancelExport();
  },

  cancelExport: function () {
    if (this.currentExportDeferred)
      this.currentExportDeferred.cancel();
  },

  exportEML: function () {
    return this.currentExportDeferred = this.exportEMLInternalDeferred(false);
  },

  exportEMLRecursive: function () {
    return this.currentExportDeferred = this.exportEMLInternalDeferred(true);
  },

  exportAllEMLInLocalFolderDeferred: function (destination, targetMessageFolder, recursive) {
    var self = this;

    var messageDirectory, errorObject;

    var exportCanceled = false;
    var exportMessagesProcess;
    var exportChildFolderProcess;
    var exportFolderProcess = Deferred.next(function () {
      return self.getMessageArrayFromFolder(targetMessageFolder); // Throws exception
    }).next(function(messages) {
      // create unique message destination directory
      messageDirectory = Util.getUniqueFile(destination, self.getMessageFolderName(targetMessageFolder), {
        isDirectory: true
      });
      messageDirectory.create(1, 0o775);
      return messages;
    }).next(function (messages) {
      exportMessagesProcess = self.saveMessagesAsEMLDeferred(messages, messageDirectory);
      return exportMessagesProcess.next(function () {
        self.folderExportingProgress.countUp();

        Util.log("recursive => " + recursive);

        if (recursive) {
          // Export sub folders recursively
          var subFolders = self.getAllSubFolders(targetMessageFolder);

          Util.log("subFolders => " + subFolders);

          var exportFailed = false;
          return Deferred.loop(subFolders.length, function (index, context) {
            // When some export failed, break this loop
            if (errorObject || exportCanceled)
              return false;

            var subFolder = subFolders[index];
            exportChildFolderProcess = self.exportAllEMLInLocalFolderDeferred(
              messageDirectory, subFolder, recursive
            ).error(function (errorFromSubCall) {
              // Failed to export a sub folder
              errorObject = errorFromSubCall;
            }).next(function (exportInformation) {
              if (exportInformation.error)
                errorObject = exportInformation.error;
            });

            return exportChildFolderProcess;
          });
        }
      });
    }).error(function (unknownError) {
      // Unknown error
      errorObject = unknownError;
    }).next(function () {
      exportMessagesProcess = null;

      return {
        createdDirectory: messageDirectory,
        error: errorObject
      };
    });

    exportFolderProcess.canceller = function () {
      exportCanceled = true;
      if (exportChildFolderProcess)
        exportChildFolderProcess.cancel();
      if (exportMessagesProcess)
        exportMessagesProcess.cancel();
    };

    return exportFolderProcess;
  },

// function selectSmartFolder() {
// 	var fTree = document.getElementById("folderTree");
// 	var fTreeSel = fTree.view.selection;
// 	if (fTreeSel.isSelected(fTreeSel.currentIndex))
// 		return;
// 	var rangeCount = fTree.view.selection.getRangeCount();
// 	var startIndex = {};
// 	var endIndex = {};
//         fTree.view.selection.getRangeAt(0, startIndex, endIndex);
// 	fTree.view.selection.currentIndex = startIndex.value;
// 	FolderPaneSelectionChange();
// }

  //
  // exportSmartFolderDeferred: function (destination, targetMessageFolder) {
  //   // getFolderForViewIndex というものもあるようす

  //   var totalMessages = targetMessageFolder.getTotalMessages(false);
  //   var foldername    = targetMessageFolder.name;

  //   for (let i = 0; i < totalMessages; ++i) {
  //     // this returns uri for selected folder
  //     let messageURI = this.global.gDBView.getURIForViewIndex(i);
  //     // TODO: get nsIMessage
  //   }
  // },

  // function exportSmartFolder(destination) {
  //   // To export virtual folder, it's necessary to select it really
  //   selectSmartFolder();
  //   setTimeout(function() {exportSmartFolderDelayed(targetMessageFolder);},1500);
  // }

  saveMessagesAsEMLDeferred: function (messages, destination) {
    var self = this;

    var saveMessagesCanceled = false;
    var getMessageProcess;
    var saveMessagesProcess = Deferred.next(function () {
      self.messageExportingProgress.resetTaskCount(messages.length);

      return Deferred.loop(messages.length, function (index) {
        if (saveMessagesCanceled)
          return false;

        var nextMessage = messages[index];
        var messageURI  = nextMessage.folder.getUriForMsg(nextMessage);

        getMessageProcess = self.getMessageTextFromURIDeferred(messageURI);
        return getMessageProcess.next(function (messageText) {
          self.saveMessageTextAsEML(nextMessage, messageText, destination);
          self.messageExportingProgress.countUp();
        });
      });
    });

    saveMessagesProcess.canceller = function () {
      saveMessagesCanceled = true;
      if (getMessageProcess)
        getMessageProcess.cancel();
    };

    return saveMessagesProcess;
  },

  getFileNameDatePartFromDate: function (date) {
    return Util.fillString(date.getFullYear(), 4) +
      Util.fillString(date.getMonth() + 1, 2) +
      Util.fillString(date.getDate(), 2) +
      Util.fillString(date.getHours(), 2) +
      Util.fillString(date.getMinutes(), 2) +
      Util.fillString(date.getSeconds(), 2);
  },

  createDateObjectFromMessage: function (message) {
    return new Date(message.dateInSeconds * 1000);
  },

  MAX_SUBJECT_LENGTH_IN_FILENAME: 20,

  getFileNameForMessage: function (message, subject) {
    return this.getFileNameDatePartFromDate(this.createDateObjectFromMessage(message)) +
      "_" + subject.substring(0, this.MAX_SUBJECT_LENGTH_IN_FILENAME) + ".eml";
  },

  escapeFromInMessage: function (messageText) {
    // See https://bugzilla.mozilla.org/show_bug.cgi?id=119441 and
    // https://bugzilla.mozilla.org/show_bug.cgi?id=194382
    return messageText.replace(/\nFrom /g, "\n From ").replace(/^From.+\r?\n/, "");
  },

  writeMessageFile: function (messageFile, messageText) {
    var foStream = Cc["@mozilla.org/network/file-output-stream;1"]
	  .createInstance(Ci.nsIFileOutputStream);
    // TODO: check permission (0664)
    foStream.init(messageFile, 0x02 | 0x08 | 0x20, 436, 0); // write, create, truncate

    if (messageText)
      foStream.write(messageText, messageText.length);
    foStream.close();
  },

  saveMessageTextAsEML: function (message, messageText, destination) {
    var subject = message.mime2DecodedSubject || "";
    var escapedMessageText = this.escapeFromInMessage(messageText);

    var messageFile = Util.getUniqueFile(destination, this.getFileNameForMessage(message, subject));
    messageFile.createUnique(0, 0o644);

    this.writeMessageFile(messageFile, escapedMessageText); // TODO: check encoding
  },

  getMessageTextFromURIDeferred: function (messageURI) {
    var canceled = false;
    var deferred = new Deferred();
    deferred.canceller = function () {
      canceled = true;
    };

    var readText = "";
    var listener = {
      onStartRequest: function (aRequest, aContext) {},

      onStopRequest: function (aRequest, aContext, aStatusCode) {
        if (!canceled)
          deferred.call(readText);
      },

      onDataAvailable: function (aRequest, aContext, aInputStream, aOffset, aCount) {
	var scriptStream = Cc["@mozilla.org/scriptableinputstream;1"].
          createInstance().QueryInterface(Ci.nsIScriptableInputStream);
	scriptStream.init(aInputStream);
        readText += scriptStream.read(scriptStream.available());
      },

      QueryInterface: function (iid)  {
        if (iid.equals(Ci.nsIStreamListener) || iid.equals(Ci.nsISupports))
          return this;
        throw Cr.NS_NOINTERFACE;
      }
    };

    var messenger = Cc["@mozilla.org/messenger;1"].createInstance(Ci.nsIMessenger);
    var messageService = messenger.messageServiceFromURI(messageURI)
	  .QueryInterface(Ci.nsIMsgMessageService);

    messageService.streamMessage(messageURI, listener, this.global.msgWindow, null, false, null);

    return deferred;
  }
};
