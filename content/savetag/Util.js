var Cc = Components.classes;
var Ci = Components.interfaces;
var Cu = Components.utils;
var Cr = Components.results;

var Util = {
  DEBUG: false,

  getEnv: function (aName, aDefault) {
    var env = Cc['@mozilla.org/process/environment;1']
      .getService(Ci.nsIEnvironment);

    return env.exists(aName) ?
      env.get(aName) : (1 in arguments ? arguments[1] : null);
  },

  or: function (aValue, aDefault) {
    return typeof aValue === "undefined" ? aDefault : aValue;
  },

  openFile: function (aPath) {
    let file = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsILocalFile);
    file.initWithPath(aPath);

    return file;
  },

  getFile: function (aTarget) {
    let file;
    if (aTarget instanceof Ci.nsIFile) {
      file = aTarget.clone();
    } else {
      file = Util.openFile(aTarget);
    }

    return file;
  },

  readFile: function (aTarget, aOptions) {
    aOptions = aOptions || {};

    let file = Util.getFile(aTarget);
    if (!file.exists())
      throw new Error(file.path + " not found");

    let fileStream = Cc["@mozilla.org/network/file-input-stream;1"]
      .createInstance(Ci.nsIFileInputStream);
    fileStream.init(file,
                    Util.or(aOptions.ioFlags, 1),
                    Util.or(aOptions.permission, 0),
                    Util.or(aOptions.behaviorFlags, false));

    let converter = aOptions.converter;
    if (!converter) {
      converter = Cc["@mozilla.org/intl/converter-input-stream;1"]
        .createInstance(Ci.nsIConverterInputStream);

      converter.init(fileStream,
                     Util.or(aOptions.charset, 'UTF-8'),
                     fileStream.available(),
                     converter.DEFAULT_REPLACEMENT_CHARACTER);
    }

    let out = {};
    converter.readString(fileStream.available(), out);
    fileStream.close();

    return out.value;
  },

  writeFile: function (aTarget, aData, aOptions) {
    aOptions = aOptions || {};

    let file = Util.getFile(aTarget);
    if (file.exists() && !Util.or(aOptions.overwrite, true))
      throw new Error(file.path + " already exists");

    let fileStream = Cc["@mozilla.org/network/file-output-stream;1"]
      .createInstance(Ci.nsIFileOutputStream);
    fileStream.init(file,
                    Util.or(aOptions.ioFlags, 0x02 | 0x08 | 0x20),
                    Util.or(aOptions.permission, 0644),
                    Util.or(aOptions.behaviorFlags, false));

    let wrote = fileStream.write(aData, aData.length);
    if (wrote != aData.length)
      throw new Error("Failed to write data to " + file.path);

    fileStream.close();

    return wrote;
  },

  getSpecialDirectory: function (aProp) {
    var dirService = Cc['@mozilla.org/file/directory_service;1']
      .getService(Ci.nsIProperties);

    return dirService.get(aProp, Ci.nsILocalFile);
  },

  copyDirectoryAs: function (source, dest) {
    let sourceDir = this.getFile(source);
    let destDir = this.getFile(dest);

    if (destDir.exists())
      destDir.remove(true);
    sourceDir.copyTo(destDir.parent, destDir.leafName);
  },

  readDirectory: function (directory) {
    directory = Util.getFile(directory);

    if (directory.exists() && directory.isDirectory()) {
      let entries = directory.directoryEntries;
      let array = [];
      while (entries.hasMoreElements())
        array.push(entries.getNext().QueryInterface(Ci.nsIFile));
      return array;
    }

    return null;
  },

  // Get file for file-1, file-2, file-3, ...
  // Currently, for directory only (do not consider suffixes)
  getIdenticalFileFor: function (originalFile) {
    var parent       = originalFile.parent;
    var file         = originalFile;
    var originalName = originalFile.leafName;

    var i = 0;
    while (file.exists()) {
      file = parent.clone();
      file.append(originalName + "-" + (++i));
    }

    return file;
  },

  fillString: function fillString(string, preferredLength, fillingCharacter) {
    string = "" + string;
    if (typeof fillingCharacter === "undefined")
      fillingCharacter = "0";
    return (Array(preferredLength - string.length + 1).join(fillingCharacter) + string);
  },

  format: function (formatString) {
    formatString = "" + formatString;

    var values = Array.slice(arguments, 1);

    return formatString.replace(/%s/g, function () {
      return Util.or(values.shift(), "");
    });
  },

  formatBytes: function (bytes, base) {
    base = base || 1024;

    var notations = ["", "K", "M", "G", "T", "P", "E", "Z", "Y"];
    var number = bytes;

    while (number >= base && notations.length > 1) {
      number /= base;
      notations.shift();
    }

    return [number.toFixed(0), notations[0] + "B"];
  },

  splitFileNameWithExtension: function (filename) {
    if (filename.indexOf(".") < 0)
      return [filename, null];
    // var matched = filename.match(/^([^.]*?)((?:\..*)*)$/);
    var matched = filename.match(/^(.*)(\.[^.]*)$/);
    return [matched[1], matched[2]];
  },

  extractExtensionFromFileName: function extractExtensionFromFileName (filename) {
    return this.splitFileNameWithExtension(filename)[1];
  },

  removeExtensionFromFileName: function removeExtension (filename) {
    return this.splitFileNameWithExtension(filename)[0];
  },

  getUniqueFile: function (directory, rawBasename, options) {
    options = options || {};
    var isDirectory = options.isDirectory;

    function getFile(name) {
      var file = directory.clone();
      file.append(name);
      return file;
    }

    var basename = Util.escapeFileName(rawBasename);

    var basenameWithoutExtension, extension;
    if (isDirectory) {
      basenameWithoutExtension = basename;
      extension = "";
    } else {
      basenameWithoutExtension = Util.removeExtensionFromFileName(basename);
      extension = Util.extractExtensionFromFileName(basename) || "";
    }

    var file = getFile(basename);
    var number = 1;             // file(1), file(2), ...
    while (file.exists()) {
      var fileName = basenameWithoutExtension + "(" + number++ + ")" + extension;
      file = getFile(fileName);
    }

    return file;
  },

  escapeFileName: function (filename, replacer) {
    if (typeof replacer === "undefined")
      replacer = "_";
    return filename.replace(/[\x00-\x19]/g, replacer).replace(/[\/\\:,<>*\?\"\|]/g, replacer);
  },

  log: function () {
    if (this.DEBUG) {
      var consoleService = Cc["@mozilla.org/consoleservice;1"].getService(Ci.nsIConsoleService);
      consoleService.logStringMessage(Util.format.apply(Util, arguments));
    }
  },

  generateUUID: function () {
    return Cc["@mozilla.org/uuid-generator;1"]
      .getService(Ci.nsIUUIDGenerator)
      .generateUUID();
  },

  openFileFromURL: function (urlSpec) {
    let ios = Cc['@mozilla.org/network/io-service;1']
          .getService(Ci.nsIIOService);
    var fileHandler = ios.getProtocolHandler('file')
          .QueryInterface(Ci.nsIFileProtocolHandler);
    return fileHandler.getFileFromURLSpec(urlSpec);
  },

  chromeToURLSpec: function (aUrl) {
    if (!aUrl || !(/^chrome:/.test(aUrl)))
      return null;

    let uri = this.makeURIFromSpec(aUrl);
    let cr = Cc['@mozilla.org/chrome/chrome-registry;1']
          .getService(Ci.nsIChromeRegistry);
    let urlSpec = cr.convertChromeURL(uri).spec;

    return urlSpec;
  },

  makeURIFromSpec: function (aURI) {
    let ios = Cc['@mozilla.org/network/io-service;1']
          .getService(Ci.nsIIOService);
    return ios.newURI(aURI, 'UTF-8', null);
  },

  launchProcess: function (exe, args) {
    let process = Cc["@mozilla.org/process/util;1"].createInstance(Ci.nsIProcess);
    let exe = Util.getFile(exe);
    process.init(exe);
    process.run(false, args, args.length);
    return process;
  },

  restartApplication: function () {
    const nsIAppStartup = Ci.nsIAppStartup;

    let os         = Cc["@mozilla.org/observer-service;1"].getService(Ci.nsIObserverService);
    let cancelQuit = Cc["@mozilla.org/supports-PRBool;1"].createInstance(Ci.nsISupportsPRBool);

    os.notifyObservers(cancelQuit, "quit-application-requested", null);
    if (cancelQuit.data)
      return;

    os.notifyObservers(null, "quit-application-granted", null);
    let wm = Cc["@mozilla.org/appshell/window-mediator;1"].getService(Ci.nsIWindowMediator);
    let windows = wm.getEnumerator(null);

    while (windows.hasMoreElements()) {
      let win = windows.getNext();
      if (("tryToClose" in win) && !win.tryToClose())
        return;
    }

    Cc["@mozilla.org/toolkit/app-startup;1"].getService(nsIAppStartup)
      .quit(nsIAppStartup.eRestart | nsIAppStartup.eAttemptQuit);
  },

  getMainWindow: function () {
    return Cc["@mozilla.org/appshell/window-mediator;1"]
      .getService(Ci.nsIWindowMediator)
      .getMostRecentWindow("mail:3pane");
  },

  openDialog: function (owner, url, name, features, arguments) {
    let windowWatcher = Cc["@mozilla.org/embedcomp/window-watcher;1"]
      .getService(Ci.nsIWindowWatcher);

    if (arguments !== undefined && arguments !== null) {
      let array = Cc["@mozilla.org/supports-array;1"]
                    .createInstance(Ci.nsISupportsArray);
      arguments.forEach(function(aItem) {
        if (aItem === null ||
          aItem === void(0) ||
          aItem instanceof Ci.nsISupports) {
          array.AppendElement(aItem);
        } else {
          let variant = Cc["@mozilla.org/variant;1"]
                        .createInstance(Ci.nsIVariant)
                        .QueryInterface(Ci.nsIWritableVariant);
          variant.setFromVariant(aItem);
          aItem = variant;
        }
        array.AppendElement(aItem);
      }, this);
      arguments = array;
    }

    windowWatcher.openWindow(owner || null, url, name, features, arguments || null);
  },

  // ============================================================
  // DOM
  // ============================================================

  alert: function (aTitle, aMessage, aWindow) {
    let prompts = Cc["@mozilla.org/embedcomp/prompt-service;1"]
          .getService(Ci.nsIPromptService);
    prompts.alert(aWindow || window, aTitle, aMessage);
  },

  alert2: function () {
    let message = Util.format.apply(Util, arguments);
    Util.alert("Alert", message);
  },

  confirm: function (aTitle, aMessage, aWindow) {
    return this.prompts.confirm(aWindow || window, aTitle, aMessage);
  },

  get prompts() {
    return Cc["@mozilla.org/embedcomp/prompt-service;1"]
          .getService(Ci.nsIPromptService);
  },
  confirmEx: function (parent, title, message, flags, firstbutton, secondbutton, thirdbutton, checkboxlabel, checked) {
    return this.prompts.confirmEx(
      parent || null,
      title,
      message,
      flags,
      firstbutton || null,
      secondbutton || null,
      thirdbutton || null,
      checkboxlabel || null,
      checked || {}
    );
  },

  getElementCreator: function (doc) {
    return function elementCreator(name, attrs, children) {
      let elem = doc.createElement(name);

      if (attrs)
        for (let [k, v] in Iterator(attrs))
          elem.setAttribute(k, v);

      if (children)
        for (let [k, v] in Iterator(children))
          elem.appendChild(v);

      return elem;
    };
  },

  http: {
    params:
    function params(prm) {
      let pt = typeof prm;

      if (prm && pt === "object")
        prm = [k + "=" + v for ([k, v] in Iterator(prm))].join("&");
      else if (pt !== "string")
        prm = "";

      return prm;
    },

    request:
    function request(method, url, callback, params, opts) {
      opts = opts || {};

      let req = Cc["@mozilla.org/xmlextras/xmlhttprequest;1"].createInstance();
      req.QueryInterface(Ci.nsIXMLHttpRequest);

      const async = (typeof callback === "function");

      if (async)
        req.onreadystatechange = function () { if (req.readyState === 4) callback(req); };

      req.open(method, url, async, opts.username, opts.password);

      if (opts.raw)
        req.overrideMimeType('text/plain; charset=x-user-defined');

      for (let [name, value] in Iterator(opts.header || {}))
        req.setRequestHeader(name, value);

      req.send(params || null);

      return async ? void 0 : req;
    },

    get:
    function get(url, callback, params, opts) {
      params = this.params(params);
      if (params)
        url += "?" + params;

      return this.request("GET", url, callback, null, opts);
    },

    post:
    function post(url, callback, params, opts) {
      params = this.params(params);

      opts = opts || {};
      opts.header = opts.header || {};
      opts.header["Content-type"] = "application/x-www-form-urlencoded";
      opts.header["Content-length"] = params.length;
      opts.header["Connection"] = "close";

      return this.request("POST", url, callback, params, opts);
    }
  },

  toArray: function (enumerator, iface) {
    iface = iface || Ci.nsISupports;
    let array = [];

    // See http://mxr.mozilla.org/comm-central/source/mozilla/xpcom/ds/nsIArray.idl#123
    if (enumerator instanceof Ci.nsIArray)
      enumerator = enumerator.enumerate();

    if (enumerator instanceof Ci.nsISupportsArray) {
      let count = enumerator.Count();
      for (let i = 0; i < count; ++i)
        array.push(enumerator.QueryElementAt(i, iface));
    } else if (enumerator instanceof Ci.nsISimpleEnumerator) {
      while (enumerator.hasMoreElements())
        array.push(enumerator.getNext().QueryInterface(iface));
    }

    return array;
  },

  equal: function (a, b, propNames) {
    return propNames.every(function (propName) { return a[propName] === b[propName]; });
  }
};
