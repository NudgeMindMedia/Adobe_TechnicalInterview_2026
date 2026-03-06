/* AE Render Daemon (Option A: run once, keep AE open)
   - Watches ready/ for manifests (polling via app.scheduleTask)
   - For each manifest:
       - resolves paths relative to LOCAL_DRIVE_ROOT
       - waits for file stability (size >= min, unchanged for stable seconds)
       - replaces footage by tag-or-name (stable across runs)
       - forces footage item name back to target name
       - renders comp with templates to output path
       - moves manifest to processed/ or failed/ + error marker

   Fixes in this version:
   - No undo groups (prevents "Undo group mismatch" warnings)
   - Closes project WITHOUT saving before opening template (prevents save prompts)
   - Closes template WITHOUT saving after render (keeps loop hands-off)
   - Best-effort dialog suppression

   Added in this version:
   - Auto-adjust guardrails for image inputs so it doesn't hang on <1MB files
   - Graceful stop support (daemonRunning checks + stopDaemon())
*/

(function () {
  // ---------------- CONFIG ----------------
  var LOCAL_DRIVE_ROOT = "C:/Users/User/My Drive/Adobe_TechnicalInterview_2026/AE_Automation/";
  var READY_DIR = LOCAL_DRIVE_ROOT + "ready/";
  var PROCESSED_DIR = LOCAL_DRIVE_ROOT + "processed/";
  var FAILED_DIR = LOCAL_DRIVE_ROOT + "failed/";

  var POLL_INTERVAL_MS = 2000;       // 2 seconds
  var MAX_WAIT_SYNC_SECONDS = 900;   // 15 minutes max to wait for file to stabilize

  var LOG_PATH = LOCAL_DRIVE_ROOT + "scripts/ae_daemon.log";

  // Tag to permanently identify the swap target footage item (robust even if AE renames it)
  var SWAP_TARGET_TAG = "__SWAP_TARGET__";

  // ---------------- STATE ----------------
  var daemonRunning = true;
  var isBusy = false;

  // ---------------- HELPERS ----------------
  function log(msg) {
    try {
      var f = new File(LOG_PATH);
      f.open("a");
      f.writeln(new Date().toUTCString() + " | " + msg);
      f.close();
    } catch (e) {}
  }

  function ensureFolder(path) {
    var folder = new Folder(path);
    if (!folder.exists) folder.create();
  }

  // Best-effort dialog suppression (varies by AE version)
  function suppressDialogsOn() {
    try { if (app.beginSuppressDialogs) app.beginSuppressDialogs(); } catch (e) {}
  }
  function suppressDialogsOff() {
    try { if (app.endSuppressDialogs) app.endSuppressDialogs(); } catch (e) {}
  }

  // Close current project without saving (prevents prompts)
  function closeProjectNoSave() {
    try {
      if (app.project && app.project.numItems > 0) {
        app.project.close(CloseOptions.DO_NOT_SAVE_CHANGES);
      }
    } catch (e) {
      log("WARN closeProjectNoSave failed: " + e.toString());
    }
  }

  // ExtendScript-safe absolute path detection (no regex)
  function isAbsolutePath(p) {
    if (!p) return false;
    if (p.length >= 2 && p.charAt(1) === ":") return true;
    if (p.length >= 2 && p.charAt(0) === "\\" && p.charAt(1) === "\\") return true;
    if (p.length >= 2 && p.charAt(0) === "/" && p.charAt(1) === "/") return true;
    return false;
  }

  function toAbs(p) {
    if (!p) return p;
    if (isAbsolutePath(p)) return p;
    return LOCAL_DRIVE_ROOT + p;
  }

  function readTextFile(path) {
    var f = new File(path);
    if (!f.exists) throw new Error("File not found: " + path);
    f.open("r");
    var s = f.read();
    f.close();
    return s;
  }

  function writeTextFile(path, text) {
    var f = new File(path);
    f.open("w");
    f.write(text);
    f.close();
  }

  function moveFile(fromPath, toPath) {
    var src = new File(fromPath);
    if (!src.exists) throw new Error("Cannot move missing file: " + fromPath);
    var dst = new File(toPath);

    if (dst.exists) {
      var base = dst.fullName.replace(/\.json$/i, "");
      var alt = base + "__" + (new Date().getTime()) + ".json";
      dst = new File(alt);
    }

    var ok = src.rename(dst.fullName);
    if (!ok) {
      src.copy(dst.fullName);
      src.remove();
    }
  }

  function listJsonFiles(dirPath) {
    var folder = new Folder(dirPath);
    if (!folder.exists) return [];
    var files = folder.getFiles(function (f) {
      return (f instanceof File) && (f.name.toLowerCase().indexOf(".json") === f.name.length - 5);
    });

    files.sort(function (a, b) {
      var an = a.name.toLowerCase();
      var bn = b.name.toLowerCase();
      return an < bn ? -1 : (an > bn ? 1 : 0);
    });

    return files;
  }

  // Wait until:
  // - file exists
  // - size >= minBytes
  // - size unchanged for stableSeconds
  function waitForStableFile(path, minBytes, stableSeconds) {
    var start = new Date().getTime();
    var lastSize = -1;
    var stableStart = null;

    while (true) {
      if (!daemonRunning) {
        throw new Error("Daemon stopped by user.");
      }

      var now = new Date().getTime();
      var elapsed = (now - start) / 1000;
      if (elapsed > MAX_WAIT_SYNC_SECONDS) {
        throw new Error("Timed out waiting for file to stabilize: " + path);
      }

      var f = new File(path);
      if (!f.exists) {
        $.sleep(1000);
        continue;
      }

      var size = f.length;
      if (size < minBytes) {
        $.sleep(1000);
        continue;
      }

      if (size !== lastSize) {
        lastSize = size;
        stableStart = new Date().getTime();
      } else {
        if (stableStart !== null) {
          var stableElapsed = (new Date().getTime() - stableStart) / 1000;
          if (stableElapsed >= stableSeconds) {
            return; // stable
          }
        }
      }

      $.sleep(1000);
    }
  }

  function findFootageSwapTarget(targetName) {
    // Tag-first BUT require name match when a targetName is provided
    for (var i = 1; i <= app.project.numItems; i++) {
      var item = app.project.item(i);
      if (item instanceof FootageItem) {
        try {
          if (item.comment === SWAP_TARGET_TAG) {
            if (!targetName || item.name === targetName) {
              return item;
            }
          }
        } catch (e) {}
      }
    }

    // Name fallback
    for (var j = 1; j <= app.project.numItems; j++) {
      var it = app.project.item(j);
      if (it instanceof FootageItem && it.name === targetName) {
        return it;
      }
    }

    return null;
  }

  function clearRenderQueue() {
    while (app.project.renderQueue.numItems > 0) {
      app.project.renderQueue.item(1).remove();
    }
  }

  function findCompByName(name) {
    for (var i = 1; i <= app.project.numItems; i++) {
      var item = app.project.item(i);
      if (item instanceof CompItem && item.name === name) {
        return item;
      }
    }
    return null;
  }

  function openTemplateFresh(templatePath) {
    suppressDialogsOn();
    closeProjectNoSave();

    var projFile = new File(templatePath);
    if (!projFile.exists) {
      suppressDialogsOff();
      throw new Error("AE project not found: " + templatePath);
    }

    app.open(projFile);
    suppressDialogsOff();
  }

  function renderJob(manifest, manifestPath) {
    var templatePath = toAbs(manifest.ae_project.template_path);
    var replacementMp4Path = toAbs(manifest.swap.replacement_mp4_path);

    var outputDir = toAbs(manifest.output.output_dir);
    if (outputDir.charAt(outputDir.length - 1) !== "/" && outputDir.charAt(outputDir.length - 1) !== "\\") {
      outputDir += "/";
    }
    var outputFile = outputDir + manifest.output.output_filename;

    var compName = manifest.ae_project.comp_name;
    var targetFootageName = manifest.swap.target_footage_name;

    var renderSettings = manifest.ae_project.render_settings || "Best Settings";
    var outputModule = manifest.ae_project.output_module || "H.264";

    var stableSeconds = (manifest.sync_guardrails && manifest.sync_guardrails.stable_seconds_required) ? manifest.sync_guardrails.stable_seconds_required : 8;
    var minBytes = (manifest.sync_guardrails && manifest.sync_guardrails.expected_min_bytes) ? manifest.sync_guardrails.expected_min_bytes : 1000000;

    // Auto-adjust for image inputs so we don't hang on <1MB files
    var lowerPath = String(replacementMp4Path).toLowerCase();
    var isImage = (
      lowerPath.indexOf(".png") > -1 ||
      lowerPath.indexOf(".jpg") > -1 ||
      lowerPath.indexOf(".jpeg") > -1 ||
      lowerPath.indexOf(".webp") > -1
    );

    if (isImage) {
      if (minBytes > 200000) minBytes = 200000;
      if (minBytes < 5000) minBytes = 5000;
      if (stableSeconds > 3) stableSeconds = 3;
    } else {
      if (minBytes < 1000000) minBytes = 1000000;
      if (stableSeconds < 6) stableSeconds = 6;
    }

    ensureFolder(outputDir);
    ensureFolder(PROCESSED_DIR);
    ensureFolder(FAILED_DIR);

    log("JOB START " + manifest.job_id + " | manifest=" + manifestPath);
    log("Project=" + templatePath);
    log("ASSET=" + replacementMp4Path + " | minBytes=" + minBytes + " | stableSeconds=" + stableSeconds);
    log("Output=" + outputFile);

    waitForStableFile(replacementMp4Path, minBytes, stableSeconds);

    openTemplateFresh(templatePath);

    var footage = findFootageSwapTarget(targetFootageName);
    if (!footage) throw new Error("Target footage not found (tag or name): " + targetFootageName);

    var mp4File = new File(replacementMp4Path);
    if (!mp4File.exists) throw new Error("Replacement asset missing: " + replacementMp4Path);

    footage.replace(mp4File);

    try { footage.comment = SWAP_TARGET_TAG; } catch (eTag) {}
    try { footage.name = targetFootageName; } catch (eName) {}

    var comp = findCompByName(compName);
    if (!comp) throw new Error("Comp not found: " + compName);

    clearRenderQueue();
    var rqItem = app.project.renderQueue.items.add(comp);

    try { rqItem.applyTemplate(renderSettings); }
    catch (e1) { log("WARN render settings template not applied: " + renderSettings + " | " + e1.toString()); }

    var om = rqItem.outputModule(1);
    try { om.applyTemplate(outputModule); }
    catch (e2) { throw new Error("Output module template not found/applied: " + outputModule + " | " + e2.toString()); }

    om.file = new File(outputFile);

    app.project.renderQueue.render();

    var outF = new File(outputFile);
    if (!outF.exists || outF.length <= 0) {
      throw new Error("Render finished but output missing/empty: " + outputFile);
    }

    log("JOB SUCCESS " + manifest.job_id + " | output=" + outputFile);

    suppressDialogsOn();
    closeProjectNoSave();
    suppressDialogsOff();
  }

  function processOneManifest() {
    if (!daemonRunning || isBusy) return;

    var files = listJsonFiles(READY_DIR);
    if (!files || files.length === 0) return;

    var mf = files[0];
    var mfPath = mf.fullName;

    isBusy = true;

    try {
      var raw = readTextFile(mfPath);
      var manifest = JSON.parse(raw);

      if (!manifest.job_type || manifest.job_type !== "replace_footage_and_render") {
        throw new Error("Invalid job_type: " + manifest.job_type);
      }

      renderJob(manifest, mfPath);

      var dst = PROCESSED_DIR + mf.name;
      moveFile(mfPath, dst);

    } catch (err) {
      var msg = (err && err.toString) ? err.toString() : String(err);
      log("JOB FAIL " + mf.name + " | " + msg);

      var errPath = FAILED_DIR + mf.name.replace(/\.json$/i, "") + ".error.txt";
      writeTextFile(errPath, msg);

      var dstFail = FAILED_DIR + mf.name;
      try { moveFile(mfPath, dstFail); } catch (moveErr) {}

      suppressDialogsOn();
      closeProjectNoSave();
      suppressDialogsOff();

    } finally {
      isBusy = false;
    }
  }

  function tick() {
    if (!daemonRunning) {
      log("Daemon stopped.");
      return;
    }

    try { processOneManifest(); }
    catch (e) { log("Tick error: " + e.toString()); }

    app.scheduleTask("tick()", POLL_INTERVAL_MS, false);
  }

  $.global.tick = tick;

  // Manual stop hook (type stopDaemon() in the ExtendScript console)
  $.global.stopDaemon = function () {
    daemonRunning = false;
    log("Daemon stop requested.");
  };

  ensureFolder(READY_DIR);
  ensureFolder(PROCESSED_DIR);
  ensureFolder(FAILED_DIR);
  ensureFolder(LOCAL_DRIVE_ROOT + "scripts/");

  log("AE Render Daemon started. Watching: " + READY_DIR);
  tick();

})();
