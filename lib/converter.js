var path = require("path");
var fs = require("fs");
var os = require("os");
var helper = require("./helper");
var spawn = require("child_process").spawn;
var params = require("./params");
var debug = require("debug")("carbone:converter");
var which = require("which");

/* Factories object */
var conversionFactory = {};
/* An active factory is a factory which is starting (but not started completely), running or stopping (but not stopped completely) */
var activeFactories = [];
/* Every conversion is placed in this job queue */
var jobQueue = [];
/* If true, a factory will restart automatically */
var isAutoRestartActive = true;

var isLibreOfficeFound = false;

var converterOptions = {
  /* Python path */
  pythonExecPath: "python",
  /* Libre Office executable path */
  sofficeExecPath: "soffice",
  /* Delay before killing the other process (either LibreOffice or Python) when one of them died */
  delayBeforeKill: 500,
};

/* get the total memory available on the system (unit: MB) */
const totalMemoryAvailableMB = os.totalmem() / 1024 / 1024;

var pythonErrors = {
  1: "Global error",
  100: "Existing office server not found",
  400: "Could not open document",
  401: "Could not convert document",
};

var converter = {
  /**
   * Initialize the converter.
   * @param {Object}   options : same options as carbone's options
   * @param {function} callback(factory): called when all factories are ready. if startFactory is true, the first parameter will contain the object descriptor of all factories
   */
  init: function (options, callback) {
    if (typeof options === "function") {
      callback = options;
    } else {
      for (var attr in options) {
        if (params[attr] !== undefined) {
          params[attr] = options[attr];
        } else {
          throw Error("Undefined options :" + attr);
        }
      }
    }
    // restart Factory automatically if it crashes.
    isAutoRestartActive = true;

    // if we must start all factory now
    if (params.startFactory === true) {
      // and if the maximum of factories is not reached
      if (activeFactories.length < params.factories) {
        var _nbFactoriesStarting = 0;
        for (var i = 0; i < params.factories; i++) {
          _nbFactoriesStarting++;
          addConversionFactory(function () {
            // here all factories are ready
            _nbFactoriesStarting--;
            if (_nbFactoriesStarting === 0 && callback) {
              callback(conversionFactory);
            }
          });
        }
      }
    } else {
      // else, start LibreOffice when needed
      if (callback) {
        callback();
      }
    }
  },

  /**
   * Kill all LibreOffice + Python threads
   * When this method is called, we must call init() to re-initialize the converter
   *
   * @param {function} callback : when everything is off
   */
  exit: function (callback) {
    isAutoRestartActive = false;
    jobQueue = [];
    for (var i in conversionFactory) {
      var _factory = conversionFactory[i];
      // if a factory is running
      if (
        _factory &&
        (_factory.pythonThread !== null || _factory.officeThread !== null)
      ) {
        clearTimeout(_factory.timeoutId);
        _factory.exitCallback = factoryExitFn;
        // kill Python thread first.
        if (_factory.pythonThread !== null) {
          _factory.pythonThread.kill("SIGKILL");
        }
        if (_factory.officeThread !== null) {
          _factory.officeThread.kill("SIGKILL");
          helper.rmDirRecursive(_factory.userCachePath);
        }
      }
    }
    // if all factories are already off
    if (activeFactories.length === 0) {
      factoryExitFn();
    }

    function factoryExitFn() {
      if (activeFactories.length === 0) {
        conversionFactory = {};
        debug("exit!");
        if (callback !== undefined) {
          callback();
        }
      }
    }
  },

  /**
   * Convert a document
   *
   * @param {string} inputFile : absolute path to the source document
   * @param {string} outputType : destination type of format.js (ex. writer_pdf_Export for PDF)
   * @param {string} formatOptions : options string passed to convert
   * @param {string} outputFile : outputFile to generate
   * @param {function} callback(err, outputFile)
   */
  convertFile: function (
    inputFile,
    outputType,
    formatOptions,
    outputFile,
    callback
  ) {
    if (isLibreOfficeFound === false) {
      return callback(
        "Cannot find LibreOffice. Document conversion cannot be used"
      );
    }

    var _job = {
      inputFilePath: inputFile,
      outputFilePath: outputFile,
      outputFormat: outputType,
      formatOptions: formatOptions || "",
      callback: callback,
      nbAttempt: 0,
      error: null,
    };
    jobQueue.push(_job);
    executeQueue();
  },

  /**
   * Do we need to restart LibreOffice?
   *
   * Temporal fix for memory leaks of LibreOffice 6+
   *
   * @param  {Objecct} params
   * @param  {Integer} availableMemory system available memory
   * @param  {Integer} nbReports       nb reborts computed by one factory
   * @return {Boolean}                 true if LibreOffice must be restarted, false otherwise
   */
  shouldTheFactoryBeRestarted: function (params, availableMemory, nbReports) {
    const _percentageFactoryMemoryLoaded =
      (nbReports * params.factoryMemoryFileSize * 100) / availableMemory;
    if (
      _percentageFactoryMemoryLoaded < params.factoryMemoryThreshold ||
      params.factoryMemoryThreshold === 0
    ) {
      return false;
    }
    return true;
  },
};

/** ***************************************************************************************************************/
/* Private methods */
/** ***************************************************************************************************************/

/**
 * Add a LibreOffice + Python factory (= 2 threads)
 *
 * WARNING: the callback must be used only by converter.init()
 *
 * @param {function} callback : function() called when the factory is ready to convert documents.
 */
function addConversionFactory(callback) {
  debug("ask to add a conversion factory");
  // find a free factory
  var _prevFactory = {};
  var _startListenerID = -1;
  for (var i = 0; i < params.factories; i++) {
    _prevFactory = conversionFactory[i];
    if (_prevFactory === undefined) {
      _startListenerID = i;
      break;
    } else if (
      _prevFactory.pythonThread === null &&
      _prevFactory.officeThread === null
    ) {
      _startListenerID = i;
      break;
    }
  }
  // maximum of factories reached
  if (_startListenerID === -1) {
    if (callback) {
      callback();
    }
    return;
  }
  var _uniqueName = helper.getUID();

  // generate a unique path to a fake user profile. We cannot start multiple instances of LibreOffice if it uses the same user cache
  var _userCachePath = path.join(params.tempPath, "_office_" + _uniqueName);
  if (_prevFactory) {
    // re-use previous directory if possible (faster restart)
    if (_prevFactory.userCachePath !== undefined) {
      _userCachePath = _prevFactory.userCachePath;
    }
    // If soffice crashes as soon as it was started, the callback of the previous starting process must be passed to the new started factory
    // On Linux, it happens when LibreOffice creates its directory for the first time (oosplash seems to hide this)
    if (_prevFactory.readyCallback) {
      callback = _prevFactory.readyCallback;
    }
  }
  // generate a URL in LibreOffice's format so that it's portable across OSes:
  // see: https://wiki.openoffice.org/wiki/URL_Basics
  var _userCacheURL = convertToURL(_userCachePath);

  // generate a unique pipe name
  var _pipeName = params.pipeNamePrefix + "_" + _uniqueName;
  var _connectionString =
    "pipe,name=" + _pipeName + ";urp;StarOffice.ComponentContext";
  var _officeParams = [
    "--headless",
    "--invisible",
    "--nocrashreport",
    "--nodefault",
    "--nologo",
    "--nofirststartwizard",
    "--norestore",
    "--quickstart",
    "--nolockcheck",
    "--accept=" + _connectionString,
    "-env:UserInstallation=" + _userCacheURL,
  ];

  // save unique name
  activeFactories.push(_pipeName);

  console.log("Starting LibreOffice with params:", _officeParams);

  var _officeThread = spawn(converterOptions.sofficeExecPath, _officeParams);
  _officeThread.on(
    "close",
    generateOnExitCallback(_startListenerID, false, _pipeName)
  );
  debug("office thread started with PID " + _officeThread.pid);

  var _pythonThread = spawn(converterOptions.pythonExecPath, [
    params.pythonPath,
    "--pipe",
    _pipeName,
  ]);
  debug("python thread started with PID " + _pythonThread.pid);
  _pythonThread.on(
    "close",
    generateOnExitCallback(_startListenerID, true, _pipeName)
  );
  _pythonThread.stdout.on("data", generateOnDataCallback(_startListenerID));
  _pythonThread.stderr.on("data", function (err) {
    debug("python stderr :", err.toString());
  });

  if (_officeThread !== null && _pythonThread !== null) {
    var _factory = {
      mode: "pipe",
      pipeName: _pipeName,
      userCachePath: _userCachePath,
      pid: _officeThread.pid,
      officeThread: _officeThread,
      pythonThread: _pythonThread,
      isReady: false,
      isConverting: false,
      readyCallback: callback,
      nbrReports: 0,
      timeoutId: null,
    };
    conversionFactory[_startListenerID] = _factory;
  } else {
    throw new Error("Carbone: Cannot start LibreOffice or Python Thread");
  }
}

/**
 * Kill one LibreOffice factory
 *
 * @param  {Object} factory
 */
function killFactory(factory) {
  if (factory.isReady === false) {
    return;
  }
  factory.isReady = false;
  factory.isConverting = false;
  factory.nbrReports = 0;
  clearTimeout(factory.timeoutId);
  if (factory.officeThread !== null) {
    factory.officeThread.kill("SIGKILL");
  } else if (factory.pythonThread !== null) {
    factory.pythonThread.kill("SIGKILL");
  }
}

/**
 * Generate a callback which is used to handle thread error and exit
 * @param  {Integer} factoryID         factoryID
 * @param  {Boolean} isPythonProcess   true if the callback is used by the Python thread, false if it used by the Office Thread
 * @param  {String}  factoryUniqueName factory unique name (equals pipeName)
 * @return {Function}                  function(error)
 */
function generateOnExitCallback(factoryID, isPythonProcess, factoryUniqueName) {
  return function (error) {
    var _processName = "";
    var _otherThreadToKill = null;

    // get factory object
    var _factory = conversionFactory[factoryID];
    if (!_factory) {
      throw new Error("Carbone: Process crashed but the factory is unknown!");
    }

    // the factory cannot receive jobs anymore
    _factory.isReady = false;
    _factory.isConverting = false;
    clearTimeout(_factory.timeoutId);

    // if the Python process died...
    if (isPythonProcess === true) {
      _processName = "Python";
      _factory.pythonThread = null;
      _otherThreadToKill = _factory.officeThread;
    } else {
      _processName = "Office";
      _factory.officeThread = null;
      _otherThreadToKill = _factory.pythonThread;
    }

    debug(
      "process " +
        _processName +
        " (PID " +
        _factory.pid +
        ") of factory " +
        factoryID +
        " died " +
        error
    );

    // if both processes Python and Office are off...
    if (_factory.pythonThread === null && _factory.officeThread === null) {
      debug("factory " + factoryID + " is completely off");
      // remove factory from activeFactories to avoid infinite loop
      activeFactories.splice(activeFactories.indexOf(factoryUniqueName), 1);
      whenFactoryIsCompletelyOff(_factory);
    } else {
      _otherThreadToKill.kill("SIGKILL");
      // Fixes #12
      // SIGKILL to make sure everything is off
      //
      // Be careful, LibreOffice has two threads oosplash (parent) -> soffice (child) if launched with "soffice".
      // On Linux, we decided to launch the child process directly to simplify the thread management.
      // Otherwise, only oosplash is killed if SIGKILL is sent. In that case:
      //   - The child "soffice" is still alive and stdin, stdout and stderr of are not closed automatically.
      //   - The event "spawn.close" is received only if stdin, stdout and stderr are closed. So carbone hangs indefinitely :(
      // When killing the oospash parent process, we should close stdin, stdout and stderr and kill the child thread soffice ourself (like "pkill soffice")
      //
      // Also, we could use SIGTERM. In that case, oosplash (parent) sends the signal to its child... but this signal is not powerful enough to
      // guarantee a shutdown. If LibreOffice hangs, we could wait forever.
      //
      // It is easier to only launch directly soffice.bin directly on Linux (see below)
    }
  };
}

/**
 * Manage factory restart ot shutdown when a factory is completly off
 * @param  {Object} factory factory description
 */
function whenFactoryIsCompletelyOff(factory) {
  // if Carbone is not shutting down
  if (isAutoRestartActive === true) {
    if (factory.currentJob) {
      // if there is an error while converting a document, let's try another time
      factory.currentJob.error = new Error("Could not convert the document");
    }
    onCurrentJobEnd(factory);
    // avoid restarting too early
    setTimeout(addConversionFactory, 50);
  }
  // else if Carbone is shutting down and there is an exitCallback
  else {
    // TODO delete async
    // delete office files synchronously (we do not care because Carbone is shutting down) when office is dead
    helper.rmDirRecursive(factory.userCachePath);
    if (factory.exitCallback) {
      factory.exitCallback();
      factory.exitCallback = null;
    }
  }
}

/**
 * Generate a callback which handle communication with the Python thread
 * @param  {Integer} factoryID factoryID
 * @return {Function}          function(data)
 */
function generateOnDataCallback(factoryID) {
  return function (data) {
    var _factory = conversionFactory[factoryID];
    data = data.toString();
    // Ready to receive document conversion
    if (data === "204") {
      debug("factory " + factoryID + " ready");
      _factory.isReady = true;
      if (_factory.readyCallback) {
        _factory.readyCallback();
        // void readyCallback to avoid calling it twice when the factory object is re-used.
        _factory.readyCallback = null;
      }
      return executeQueue();
    }
    // Document converted with or without errors
    if (_factory.currentJob) {
      _factory.currentJob.error =
        pythonErrors[data] !== undefined ? new Error(pythonErrors[data]) : null;
    }
    onCurrentJobEnd(_factory);
  };
}

/**
 * Called when the job is finished
 *
 * @param  {Object} factory factory object
 */
function onCurrentJobEnd(factory) {
  var _job = factory.currentJob;
  factory.currentJob = null;
  factory.isConverting = false;
  clearTimeout(factory.timeoutId);
  if (_job && _job.callback instanceof Function) {
    // save the number of report converted to check the memory level of the LO process
    // if it reach a threshold, the LO process is killed
    if (
      converter.shouldTheFactoryBeRestarted(
        params,
        totalMemoryAvailableMB,
        ++factory.nbrReports
      ) === true
    ) {
      killFactory(factory);
    }
    _job.callback(_job.error, _job.outputFilePath);
  }
  executeQueue();
}

/**
 * Execute the queue of conversion.
 * It will auto executes itself until the queue is empty
 */
function executeQueue() {
  if (jobQueue.length === 0) {
    return;
  }
  // if there is no active factories, start them
  if (activeFactories.length < params.factories) {
    addConversionFactory();
    return;
  }
  for (var i in conversionFactory) {
    if (jobQueue.length > 0) {
      var _factory = conversionFactory[i];
      if (_factory.isReady === true && _factory.isConverting === false) {
        var _job = jobQueue.shift();
        sendToFactory(_factory, _job);
      }
    }
  }
}

/**
 * Send the document to the Factory
 *
 * @param {object} factory : LibreOffice + Python factory to send to
 * @param {object} job : job description (file to convert, callback to call when finished, ...)
 */
function sendToFactory(factory, job) {
  factory.isConverting = true;
  factory.currentJob = job;
  factory.pythonThread.stdin.write(
    '--format="' +
      job.outputFormat +
      '" --input="' +
      job.inputFilePath +
      '" --output="' +
      job.outputFilePath +
      '" --formatOptions="' +
      job.formatOptions +
      '"\n'
  );
  // keep the number of attempts to convert this file
  job.nbAttempt++;
  // Timeout to kill long conversions
  if (params.converterFactoryTimeout > 0) {
    clearTimeout(factory.timeoutId); // by security
    factory.timeoutId = setTimeout(function () {
      job.nbAttempt = params.attempts; // do not retry
      job.error = new Error(
        "Document conversion timeout reached (" +
          params.converterFactoryTimeout +
          " ms)"
      );
      killFactory(factory);
      onCurrentJobEnd(factory);
    }, params.converterFactoryTimeout);
  }
}

/**
 * Error for path
 *
 * @param {[type]} message [description]
 */
function PathError(message) {
  this.name = "PathError";
  this.code = "PathError";
  this.message = message || "Failed to convert path";
  if (typeof Error.captureStackTrace === "function") {
    Error.captureStackTrace(this, PathError);
  }
}
PathError.prototype = new Error();

/**
 * Convert an absolute path to an absolute URL understood by LibreOffice and
 *  OpenOffice. This is necessary because LO/OO use a cross-platform path format
 *  that does not match paths understood natively by OSes.
 * If the input is already a URL, it is returned as-is.
 *
 * @param {string} inputPath - An absolute path to convert to a URL.
 * @returns {string} A string suitable for use with LibreOffice as an absolute file path URL.
 */
function convertToURL(inputPath) {
  // Guard clause: if it already looks like a URL, keep it that way.
  if (inputPath.slice(0, 8) === "file:///") {
    return inputPath;
  }
  if (!path.isAbsolute(inputPath)) {
    throw new PathError("Paths to convert must be absolute");
  }
  // Split into parts so that we can join into a URL:
  var _normalizedPath = path.normalize(inputPath);
  // (Use both delimiters blindly - we're aiming for maximum compatibility)
  var _pathComponents = _normalizedPath.split(/[\\/]/);
  // Make sure there is no leading empty element, since we always add a leading "/" anyway.
  if (_pathComponents[0] === "") {
    _pathComponents.shift();
  }
  var outputURL = "file:///" + _pathComponents.join("/");
  return outputURL;
}

/**
 * Detect If LibreOffice and python are available at startup
 */
function detectLibreOffice(additionalPaths) {
  function _findBundledPython(sofficePath, pythonName) {
    debug(
      "Looking for bundled Python. sofficePath:",
      sofficePath,
      "pythonName:",
      pythonName
    );
    if (!sofficePath) {
      debug("No soffice path provided, skipping bundled Python search");
      return null;
    }

    // Try finding Python binary alongside soffice
    var _sofficeActualDirectory;
    var _symlinkDestination;
    try {
      debug("Checking if soffice path is a symlink");
      _symlinkDestination = path.resolve(
        path.dirname(sofficePath),
        fs.readlinkSync(sofficePath)
      );
      _sofficeActualDirectory = path.dirname(_symlinkDestination);
      debug("Symlink found. Actual directory:", _sofficeActualDirectory);
    } catch (error) {
      debug("Not a symlink:", error.message);
      _sofficeActualDirectory = path.dirname(sofficePath);
    }

    // Check for Python binary
    try {
      debug("Searching for Python in:", _sofficeActualDirectory);
      const pythonPath = which.sync(pythonName, {
        path: _sofficeActualDirectory,
      });
      debug("Found bundled Python at:", pythonPath);
      return pythonPath;
    } catch (error) {
      debug("No bundled Python found:", error.message);
      return null;
    }
  }

  function _findBinaries(paths, pythonName, sofficeName) {
    debug("Searching for binaries with paths:", paths);
    debug("Looking for Python:", pythonName);
    debug("Looking for LibreOffice:", sofficeName);

    var _whichSoffice;
    try {
      _whichSoffice = which.sync(sofficeName, {
        path: paths.join(":"),
        nothrow: true,
      });
      debug("Found soffice in specified paths:", _whichSoffice);
      if (!_whichSoffice) {
        _whichSoffice = which.sync(sofficeName, { nothrow: true });
        debug("Found soffice in system PATH:", _whichSoffice);
      }
    } catch (error) {
      debug("Error finding soffice:", error.message);
    }

    var _whichPython;
    try {
      _whichPython =
        _findBundledPython(_whichSoffice, "python3") ||
        _findBundledPython(_whichSoffice, "python");

      if (!_whichPython && paths.length > 0) {
        debug("Trying to find Python in specified paths");
        _whichPython =
          which.sync("python3", { path: paths.join(":"), nothrow: true }) ||
          which.sync("python", { path: paths.join(":"), nothrow: true });
        debug("Found Python in specified paths:", _whichPython);
      }

      if (!_whichPython) {
        debug("Trying to find Python in system PATH");
        _whichPython =
          which.sync("python3", { nothrow: true }) ||
          which.sync("python", { nothrow: true });
        debug("Found Python in system PATH:", _whichPython);
      }
    } catch (error) {
      debug("Error finding Python:", error.message);
    }

    return {
      soffice: _whichSoffice || null,
      python: _whichPython || null,
    };
  }

  function _listProgramDirectories(basePath, pattern) {
    try {
      return fs
        .readdirSync(basePath)
        .filter(function _isLibreOfficeDirectory(dirname) {
          return pattern.test(dirname);
        })
        .map(function _buildFullProgramPath(dirname) {
          return path.join(basePath, dirname, "program");
        });
    } catch (errorToIgnore) {
      return [];
    }
  }

  var _pathsToCheck = additionalPaths || [];
  // overridable file names to look for in the checked paths:
  var _pythonName = "python";
  var _sofficeName = "soffice";
  var _linuxDirnamePattern = /^libreoffice\d+\.\d+$/;
  var _windowsDirnamePattern = /^LibreOffice( \d+(?:\.\d+)*?)?$/i;

  if (process.platform === "darwin") {
    _pathsToCheck = _pathsToCheck.concat([
      // It is better to use the python bundled with LibreOffice:
      "/Applications/LibreOffice.app/Contents/MacOS",
      "/Applications/LibreOffice.app/Contents/Resources",
    ]);
  } else if (process.platform === "linux") {
    // on Linux, avoid oosplash parent process to simplify SIGKILL propagation. Launch directly soffice.bin.
    // Fixes #12
    _sofficeName = "soffice.bin";
    // The Document Foundation packages (.debs, at least) install to /opt,
    // into a directory named after the contained LibreOffice version.
    // Add any existing directories that match this to the list.
    _pathsToCheck = _pathsToCheck.concat(
      _listProgramDirectories("/opt", _linuxDirnamePattern)
    );
  } else if (process.platform === "win32") {
    _pathsToCheck = _pathsToCheck
      .concat(
        _listProgramDirectories("C:\\Program Files", _windowsDirnamePattern)
      )
      .concat(
        _listProgramDirectories(
          "C:\\Program Files (x86)",
          _windowsDirnamePattern
        )
      );
    _pythonName = "python.exe";
  } else {
    debug('your platform "%s" is not supported yet', process.platform);
  }

  // Common logic for all OSes: perform the search and save results as options:
  var _foundPaths = _findBinaries(_pathsToCheck, _pythonName, _sofficeName);
  if (_foundPaths.soffice) {
    debug(
      "LibreOffice found: soffice at %s, python at %s",
      _foundPaths.soffice,
      _foundPaths.python
    );
    isLibreOfficeFound = true;
    converterOptions.pythonExecPath = _foundPaths.python;
    converterOptions.sofficeExecPath = _foundPaths.soffice;
  }

  if (isLibreOfficeFound === false) {
    debug("cannot find LibreOffice. Document conversion cannot be used");
  }

  // After finding paths, let's verify the binaries work
  if (_foundPaths.soffice && _foundPaths.python) {
    debug("Verifying Python executable...");
    try {
      const pythonTest = spawn(_foundPaths.python, ["--version"]);
      pythonTest.on("error", (err) => {
        debug("Python verification failed:", err.message);
      });
      pythonTest.on("close", (code) => {
        debug("Python verification exit code:", code);
      });
    } catch (error) {
      debug("Failed to spawn Python process:", error.message);
    }

    debug("Verifying LibreOffice executable...");
    try {
      const officeTest = spawn(_foundPaths.soffice, ["--version"]);
      officeTest.on("error", (err) => {
        debug("LibreOffice verification failed:", err.message);
      });
      officeTest.on("close", (code) => {
        debug("LibreOffice verification exit code:", code);
      });
    } catch (error) {
      debug("Failed to spawn LibreOffice process:", error.message);
    }
  }
}

detectLibreOffice();

["SIGINT", "SIGHUP", "SIGQUIT"].forEach(function (signal) {
  process.on(signal, function () {
    converter.exit();
  });
});

process.on("exit", function () {
  converter.exit();
});

module.exports = converter;
