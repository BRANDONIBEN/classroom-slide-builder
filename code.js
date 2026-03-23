// ============================================================
// Classroom Slide Builder — code.js (runs in Figma sandbox)
// Supports: Essential Theology, Scripture Narrative, Biblical Finances
//
// CHANGELOG:
// 2026-03-18  Remove Dropbox, GitHub-only data. Add structured JSON support.
//             Per-session error handling in buildAllSessions. Skip PDF page
//             numbers. Increase yield time between sessions. Diagnostic logging.
// ============================================================

figma.showUI(__html__, { width: 480, height: 560 });

// Clean up any stale preview frames from previous sessions
(function cleanupStalePreview() {
  var prev = figma.currentPage.findOne(function (n) {
    return n.type === 'FRAME' && n.name === '[PREVIEW]';
  });
  if (prev) prev.remove();
})();

// Check if a slide frame is selected on launch
(function detectSelection() {
  var sel = figma.currentPage.selection;
  if (sel.length === 1 && sel[0].type === 'FRAME' && sel[0].width === 1920 && sel[0].height === 1080) {
    var frame = sel[0];
    var nameMatch = frame.name.match(/^(?:\u2691\s*)?\[(\w+)\]\s*S(\d+)\s*\u00B7\s*(\d+)\s*\u2014\s*(.*)/);
    if (nameMatch) {
      // Extract text content from the frame's text nodes
      var texts = [];
      frame.findAll(function (n) { return n.type === 'TEXT'; }).forEach(function (t) {
        texts.push({ chars: t.characters, fontSize: t.fontSize, y: t.y, x: t.x });
      });
      texts.sort(function (a, b) { return a.y - b.y; });

      var slideType = nameMatch[1].toLowerCase();
      var sessionNum = parseInt(nameMatch[2]);
      var slideNum = parseInt(nameMatch[3]);
      var titleFromName = nameMatch[4];

      var title = '';
      var body = '';
      var attribution = '';
      var bodyFontSize = 0;
      var courseName = '', sessionLabel = '';

      for (var i = 0; i < texts.length; i++) {
        var t = texts[i];
        if (t.y > 950) {
          if (t.x < 400) courseName = t.chars;
          else if (t.x > 1200) sessionLabel = t.chars;
          continue;
        }
        if (t.fontSize <= 32 && t.y < 200 && !title && t.chars !== '') {
          title = t.chars; continue;
        }
        if (t.fontSize >= 32 && !body) {
          body = t.chars; bodyFontSize = t.fontSize; continue;
        }
        if (body && !attribution && t.chars !== '' && t.fontSize < bodyFontSize && t.y > 200) {
          attribution = t.chars; continue;
        }
      }

      if (!title && titleFromName !== 'Untitled') title = titleFromName;

      var existing = getSlideOverride(sessionNum, slideNum);

      figma.ui.postMessage({
        type: 'editSlide',
        slideType: existing ? existing.type : slideType,
        sessionNum: sessionNum,
        slideNum: slideNum,
        title: existing ? existing.title : title,
        body: existing ? existing.body : body,
        attribution: existing ? (existing.attribution || '') : attribution,
        frameId: frame.id,
        frameX: frame.x,
        frameY: frame.y,
        hasOverride: !!existing,
        overrideCount: getOverrideCount(),
        versions: getSlideVersions(sessionNum, slideNum),
        courseName: courseName,
        sessionLabel: sessionLabel,
        noteData: getSlideNote(sessionNum, slideNum),
        noteCount: getNoteCount(),
        isFlagged: isSlideReviewFlagged(sessionNum, slideNum),
        flagCount: getReviewFlags().length
      });
    }
  }
  // Always send override count and page charts on launch
  setTimeout(function () {
    figma.ui.postMessage({ type: 'overrideCount', count: getOverrideCount() });
    sendPageCharts();
    // Check if print pages already exist
    var hasPrint = figma.root.children.some(function (p) { return p.name.indexOf('[PRINT]') === 0; });
    if (hasPrint) figma.ui.postMessage({ type: 'printPagesReady' });
  }, 100);
})();

// Live selection change — update edit panel when user selects a different frame
function sendSelectionData() {
  var sel = figma.currentPage.selection;

  // If the preview frame itself is selected, don't change edit state — user is editing the preview
  if (sel.length === 1 && sel[0].type === 'FRAME' && sel[0].name === '[PREVIEW]') return;

  if (sel.length === 1 && sel[0].type === 'FRAME' && sel[0].width === 1920 && sel[0].height === 1080) {
    var frame = sel[0];
    var nameMatch = frame.name.match(/^(?:\u2691\s*)?\[(\w+)\]\s*S(\d+)\s*\u00B7\s*(\d+)\s*\u2014\s*(.*)/);
    if (nameMatch) {
      var texts = [];
      frame.findAll(function (n) { return n.type === 'TEXT'; }).forEach(function (t) {
        var isItalic = false;
        try { isItalic = t.fontName && t.fontName.style && t.fontName.style.toLowerCase().indexOf('italic') !== -1; } catch (e) {}
        // Handle mixed fontSize — use getRangeFontSize for first char as fallback
        var fs = t.fontSize;
        if (typeof fs !== 'number') { try { fs = t.getRangeFontSize(0, 1); } catch (e) { fs = 0; } }
        texts.push({ chars: t.characters, fontSize: fs, y: t.y, x: t.x, isItalic: isItalic, opacity: t.opacity });
      });
      texts.sort(function (a, b) { return a.y - b.y; });
      var slideType = nameMatch[1].toLowerCase();
      var sessionNum = parseInt(nameMatch[2]);
      var slideNum = parseInt(nameMatch[3]);
      var titleFromName = nameMatch[4];
      var title = '', body = '', attribution = '';
      var bodyFontSize = 0;
      var courseName = '', sessionLabel = '';
      for (var i = 0; i < texts.length; i++) {
        var t = texts[i];
        // Footer text (below y=950) and page numbers
        if (t.y > 950) {
          if (t.x < 400) courseName = t.chars;
          else if (t.x > 1200) sessionLabel = t.chars;
          continue;
        }
        // Title: small text near top of frame
        if (t.fontSize <= 32 && t.y < 200 && !title && t.chars !== '') { title = t.chars; continue; }
        // Body: largest text in main content area
        if (t.fontSize >= 32 && !body) { body = t.chars; bodyFontSize = t.fontSize; continue; }
        // Attribution: any text after body that is smaller than body (quote/scripture slides)
        if (body && !attribution && t.chars !== '' && t.fontSize < bodyFontSize && t.y > 200) { attribution = t.chars; continue; }
      }
      if (!title && titleFromName !== 'Untitled') title = titleFromName;
      // Debug: log extraction results
      console.log('[Selection] texts:', texts.map(function(t) { return { chars: t.chars.substring(0, 40), fontSize: t.fontSize, y: Math.round(t.y), italic: t.isItalic }; }));
      console.log('[Selection] extracted:', { title: title, body: body ? body.substring(0, 40) + '...' : '', attribution: attribution, bodyFontSize: bodyFontSize });
      var existing = getSlideOverride(sessionNum, slideNum);
      // For design overrides, the override stores a frame clone, not text — use extracted text
      var useOverrideText = existing && existing.mode !== 'design' && existing.title;
      figma.ui.postMessage({
        type: 'editSlide',
        slideType: (useOverrideText ? existing.type : null) || slideType || 'body',
        sessionNum: sessionNum,
        slideNum: slideNum,
        title: useOverrideText ? existing.title : title,
        body: useOverrideText ? existing.body : body,
        attribution: useOverrideText ? (existing.attribution || '') : attribution,
        frameId: frame.id,
        frameX: frame.x,
        frameY: frame.y,
        hasOverride: !!existing,
        overrideCount: getOverrideCount(),
        versions: getSlideVersions(sessionNum, slideNum),
        courseName: courseName,
        sessionLabel: sessionLabel,
        noteData: getSlideNote(sessionNum, slideNum),
        noteCount: getNoteCount(),
        isFlagged: isSlideReviewFlagged(sessionNum, slideNum),
        flagCount: getReviewFlags().length
      });
      return;
    }
  }
  // Non-slide selected or nothing selected — do NOT remove preview or clear edit state.
  // Preview persists so user can click away to grab designs and paste onto preview.
}

figma.on('selectionchange', function () {
  sendSelectionData();
});

// Scan current page for chart/graphic frames and send to UI
function sendPageCharts() {
  var charts = [];
  figma.currentPage.children.forEach(function (node) {
    if (node.type === 'FRAME' && node.width === 1920 && node.height === 1080) {
      var m = node.name.match(/^\[(CHART|GRAPHIC)\]\s*S(\d+)\s*\u00B7\s*(\d+)\s*\u2014\s*(.*)/i);
      if (m) {
        var sNum = parseInt(m[2]);
        var slNum = parseInt(m[3]);
        var override = getSlideOverride(sNum, slNum);
        charts.push({
          frameId: node.id,
          type: m[1].toLowerCase(),
          sessionNum: sNum,
          slideNum: slNum,
          title: m[4],
          hasDesignOverride: !!(override && override.mode === 'design')
        });
      }
    }
  });
  figma.ui.postMessage({ type: 'pageCharts', charts: charts });
}

// Send chart data on page change too
figma.on('currentpagechange', function () {
  sendPageCharts();
  // Reset edit state when page changes
  figma.ui.postMessage({ type: 'resetEdit' });
});

// Clean up preview frame when plugin closes
figma.on('close', function () {
  var prev = figma.currentPage.findOne(function (n) {
    return n.type === 'FRAME' && n.name === '[PREVIEW]';
  });
  if (prev) prev.remove();
});

figma.ui.onmessage = async function (msg) {
  if (msg.type === 'build') {
    try {
      if (msg.colors) applyCustomColors(msg.colors);
      if (msg.fontConfig) await applyFontConfig(msg.fontConfig);
      var logoData = msg.logoData || { svg: null, bgColor: null };
      await buildSlides(msg.slides, msg.sessionNum, msg.sessionName, msg.courseName, logoData, msg.totalSessions || 0, msg.includePageNumbers || false);
      if (msg.exportPrintView) {
        await loadAllFonts();
        var printPage = await createPrintPage(figma.currentPage);
        figma.currentPage = printPage;
        figma.ui.postMessage({ type: 'printPagesReady' });
      }
    } catch (err) {
      figma.ui.postMessage({ type: 'error', message: String(err) });
    }
  }
  if (msg.type === 'buildAll') {
    try {
      console.log('[PLUGIN] buildAll received:', msg.sessionGroups ? msg.sessionGroups.length : 0, 'groups,', msg.totalSessions, 'totalSessions');
      if (msg.sessionGroups) {
        for (var gi = 0; gi < msg.sessionGroups.length; gi++) {
          console.log('[PLUGIN]   Group', gi, ': S' + msg.sessionGroups[gi].sessionNum, '(' + msg.sessionGroups[gi].slides.length + ' slides)');
        }
      }
      if (msg.colors) applyCustomColors(msg.colors);
      if (msg.fontConfig) await applyFontConfig(msg.fontConfig);
      var logoData = msg.logoData || { svg: null, bgColor: null };
      await buildAllSessions(msg.sessionGroups, msg.courseName, logoData, msg.totalSessions || 0, msg.includePageNumbers || false);
      if (msg.exportPrintView) {
        await loadAllFonts();
        var allPages = figma.root.children.slice();
        var lastPrintPage = null;
        for (var pi = 0; pi < allPages.length; pi++) {
          if (allPages[pi].name.indexOf('[PRINT]') !== 0 && allPages[pi].name.indexOf('[TRASH]') !== 0 && allPages[pi].name.indexOf('[STORAGE]') !== 0) {
            if (msg.courseName && allPages[pi].name.indexOf(msg.courseName) === 0) {
              lastPrintPage = await createPrintPage(allPages[pi]);
            }
          }
        }
        if (lastPrintPage) {
          figma.currentPage = lastPrintPage;
        }
        figma.ui.postMessage({ type: 'printPagesReady' });
      }
    } catch (err) {
      figma.ui.postMessage({ type: 'error', message: String(err) });
    }
  }
  if (msg.type === 'updateSlide') {
    try {
      await loadAllFonts();
      if (msg.colors) applyCustomColors(msg.colors);
      if (msg.fontConfig) await applyFontConfig(msg.fontConfig);

      // Remove old preview frame
      var preview = figma.currentPage.findOne(function (n) {
        return n.type === 'FRAME' && n.name === '[PREVIEW]';
      });
      if (preview) preview.remove();

      // Save the override so future builds use this version
      saveSlideOverride(msg.sessionNum, msg.slideNum, {
        mode: 'text',
        type: msg.slideType,
        title: msg.title,
        body: msg.body,
        attribution: msg.attribution || '',
        imageRef: msg.imageRef || ''
      });

      // Find the original frame and replace it
      var oldFrame = null;
      try { oldFrame = await figma.getNodeByIdAsync(msg.frameId); } catch (e) {}
      // Fallback: find by slide name pattern if ID lookup fails
      if (!oldFrame) {
        var namePrefix = 'S' + msg.sessionNum + ' \u00B7 ' + msg.slideNum + ' \u2014';
        oldFrame = figma.currentPage.findOne(function (n) {
          return n.type === 'FRAME' && n.name.indexOf(namePrefix) !== -1 && n.name !== '[PREVIEW]';
        });
      }
      var x = oldFrame ? oldFrame.x : msg.frameX;
      var y = oldFrame ? oldFrame.y : msg.frameY;
      var parentNode = oldFrame ? oldFrame.parent : figma.currentPage;
      // Extract page number from original frame before removing
      var existingPageNum = null;
      if (oldFrame) {
        var pgNode = oldFrame.findOne(function (n) { return n.name === '[PAGE_NUM]'; });
        if (pgNode && pgNode.type === 'TEXT') existingPageNum = pgNode.characters;
      }
      if (oldFrame) oldFrame.remove();

      // Build replacement slide — rawEdit skips ensurePeriod so user controls punctuation
      var slide = {
        type: msg.slideType,
        title: msg.title,
        body: msg.body,
        attribution: msg.attribution || '',
        sessionNum: msg.sessionNum,
        number: msg.slideNum,
        _courseName: msg.courseName || '',
        _sessionLabel: msg.sessionLabel || '',
        _rawEdit: true,
        imageRef: msg.imageRef || '',
        _pageNum: existingPageNum ? parseInt(existingPageNum) : null
      };
      var newFrame = buildFrame(slide, x, y);
      parentNode.appendChild(newFrame);

      // Build a fresh preview below the new frame
      var previewSlide = JSON.parse(JSON.stringify(slide));
      previewSlide._rawEdit = true;
      var previewFrame = buildFrame(previewSlide, x, y + H + 60);
      previewFrame.name = '[PREVIEW]';
      previewFrame.opacity = 0.85;
      parentNode.appendChild(previewFrame);

      figma.currentPage.selection = [newFrame];
      figma.viewport.scrollAndZoomIntoView([newFrame]);

      var count = getOverrideCount();
      figma.ui.postMessage({
        type: 'slideUpdated',
        overrideCount: count,
        frameId: newFrame.id,
        frameX: newFrame.x,
        frameY: newFrame.y
      });
    } catch (err) {
      figma.ui.postMessage({ type: 'error', message: String(err) });
    }
  }
  if (msg.type === 'getOverrides') {
    var allOverrides = getAllOverrides();
    figma.ui.postMessage({ type: 'overridesList', overrides: allOverrides, count: getOverrideCount() });
  }
  if (msg.type === 'clearOverride') {
    clearSlideOverride(msg.sessionNum, msg.slideNum);
    figma.ui.postMessage({ type: 'overrideCleared', count: getOverrideCount() });
  }
  if (msg.type === 'clearAllOverrides') {
    clearAllOverrides();
    figma.ui.postMessage({ type: 'overrideCleared', count: 0 });
  }
  if (msg.type === 'commitSlide') {
    try {
      // Commit the PREVIEW frame as a design override, then replace the original
      var previewFrame = figma.currentPage.findOne(function (n) {
        return n.type === 'FRAME' && n.name === '[PREVIEW]';
      });
      if (!previewFrame) {
        figma.ui.postMessage({ type: 'error', message: 'No preview frame found. Make changes first.' });
        return;
      }
      var sNum = msg.sessionNum;
      var slNum = msg.slideNum;
      if (!sNum || !slNum) {
        figma.ui.postMessage({ type: 'error', message: 'Missing slide info for commit.' });
        return;
      }

      // Give preview proper slide name before committing
      var origFrame = msg.originalFrameId ? await figma.getNodeByIdAsync(msg.originalFrameId) : null;
      var slideType = msg.slideType || 'body';
      previewFrame.name = '[' + slideType.toUpperCase() + '] S' + sNum + ' \u00B7 ' + slNum + ' \u2014 ' + (msg.title || 'Untitled');
      previewFrame.opacity = 1;

      // Commit the preview as design override
      await commitDesignOverride(sNum, slNum, previewFrame);

      // Move preview to original's position and remove original
      if (origFrame) {
        previewFrame.x = origFrame.x;
        previewFrame.y = origFrame.y;
        origFrame.remove();
      }

      figma.currentPage.selection = [previewFrame];

      // Extract text from the committed frame so the UI can rebuild
      var committedTexts = [];
      previewFrame.findAll(function (n) { return n.type === 'TEXT'; }).forEach(function (t) {
        var fs = t.fontSize;
        if (typeof fs !== 'number') { try { fs = t.getRangeFontSize(0, 1); } catch (e) { fs = 0; } }
        if (t.y < 950) committedTexts.push({ chars: t.characters, fontSize: fs, y: t.y });
      });
      committedTexts.sort(function (a, b) { return a.y - b.y; });
      var cTitle = '', cBody = '', cAttrib = '', cBodyFs = 0;
      for (var ci = 0; ci < committedTexts.length; ci++) {
        var ct = committedTexts[ci];
        if (ct.fontSize <= 32 && ct.y < 200 && !cTitle && ct.chars !== '') { cTitle = ct.chars; continue; }
        if (ct.fontSize >= 32 && !cBody) { cBody = ct.chars; cBodyFs = ct.fontSize; continue; }
        if (cBody && !cAttrib && ct.chars !== '' && ct.fontSize < cBodyFs) { cAttrib = ct.chars; continue; }
      }

      figma.ui.postMessage({
        type: 'slideCommitted',
        sessionNum: sNum,
        slideNum: slNum,
        slideType: slideType,
        frameId: previewFrame.id,
        frameX: previewFrame.x,
        frameY: previewFrame.y,
        title: cTitle,
        body: cBody,
        attribution: cAttrib,
        overrideCount: getOverrideCount()
      });
    } catch (err) {
      figma.ui.postMessage({ type: 'error', message: String(err) });
    }
  }
  if (msg.type === 'revertSlide') {
    try {
      var sel = figma.currentPage.selection;
      if (sel.length !== 1 || sel[0].type !== 'FRAME') {
        figma.ui.postMessage({ type: 'error', message: 'Select a single slide frame to revert.' });
        return;
      }
      var frame = sel[0];
      var nameMatch = frame.name.match(/^\[(\w+)\]\s*S(\d+)\s*\u00B7\s*(\d+)/);
      if (!nameMatch) {
        figma.ui.postMessage({ type: 'error', message: 'Selected frame is not a recognized slide.' });
        return;
      }
      var sNum = parseInt(nameMatch[2]);
      var slNum = parseInt(nameMatch[3]);
      // Clear override AND version history
      clearSlideOverride(sNum, slNum);
      var vKey = overrideKey(sNum, slNum) + '_versions';
      figma.root.setPluginData(vKey, '[]');
      // Extract footer from current frame before rebuild (skip page number node)
      var revertCourse = '', revertSession = '';
      var existingPageNum = null;
      frame.findAll(function (n) { return n.type === 'TEXT'; }).forEach(function (t) {
        if (t.name === '[PAGE_NUM]') { existingPageNum = parseInt(t.characters); return; }
        if (t.y > 950) {
          if (t.x < 960) revertCourse = t.characters;
          else revertSession = t.characters;
        }
      });
      // Rebuild from original source if data provided
      if (msg.originalSlide) {
        await loadAllFonts();
        if (msg.colors) applyCustomColors(msg.colors);
        if (msg.fontConfig) await applyFontConfig(msg.fontConfig);
        // Remove preview
        var prev = figma.currentPage.findOne(function (n) {
          return n.type === 'FRAME' && n.name === '[PREVIEW]';
        });
        if (prev) prev.remove();
        var x = frame.x;
        var y = frame.y;
        var parentNode = frame.parent;
        frame.remove();
        var slide = msg.originalSlide;
        slide._courseName = slide._courseName || revertCourse;
        slide._sessionLabel = slide._sessionLabel || revertSession;
        // Determine page number: from existing frame, or calculate from position if page numbers enabled
        if (existingPageNum) {
          slide._pageNum = existingPageNum;
        } else if (msg.includePageNumbers) {
          // Count slide frames to the left of this position to determine page number
          var slidesBefore = 0;
          parentNode.children.forEach(function (child) {
            if (child.type === 'FRAME' && child.width === W && child.height === H && child.x < x && child.name !== '[PREVIEW]' && !/^\[COVER\]/.test(child.name)) {
              slidesBefore++;
            }
          });
          slide._pageNum = slidesBefore + 2; // +2 for cover page offset
        }
        var newFrame = buildFrame(slide, x, y);
        parentNode.appendChild(newFrame);
        figma.currentPage.selection = [newFrame];
      }
      figma.ui.postMessage({
        type: 'slideReverted',
        sessionNum: sNum,
        slideNum: slNum,
        overrideCount: getOverrideCount(),
        restoredData: msg.originalSlide ? { type: msg.originalSlide.type, title: msg.originalSlide.title, body: msg.originalSlide.body } : null
      });
    } catch (err) {
      figma.ui.postMessage({ type: 'error', message: String(err) });
    }
  }
  if (msg.type === 'restoreVersion') {
    try {
      var sel = figma.currentPage.selection;
      if (sel.length !== 1 || sel[0].type !== 'FRAME') {
        figma.ui.postMessage({ type: 'error', message: 'Select the slide frame to restore.' });
        return;
      }
      var frame = sel[0];
      var nameMatch = frame.name.match(/^\[(\w+)\]\s*S(\d+)\s*\u00B7\s*(\d+)/);
      if (!nameMatch) return;
      var sNum = parseInt(nameMatch[2]);
      var slNum = parseInt(nameMatch[3]);
      var versions = getSlideVersions(sNum, slNum);
      var ver = versions[msg.versionIndex];
      if (!ver || !ver.data) return;
      // Extract footer from current frame before rebuilding
      var verCourse = '', verSession = '';
      frame.findAll(function (n) { return n.type === 'TEXT'; }).forEach(function (t) {
        if (t.y > 950) {
          if (t.x < 960) verCourse = t.characters;
          else verSession = t.characters;
        }
      });
      // Restore version data as current override
      var key = overrideKey(sNum, slNum);
      figma.root.setPluginData(key, JSON.stringify(ver.data));
      // Rebuild if text mode
      if (ver.data.mode === 'text') {
        await loadAllFonts();
        if (msg.colors) applyCustomColors(msg.colors);
        var x = frame.x;
        var y = frame.y;
        var parentNode = frame.parent;
        frame.remove();
        var slide = { type: ver.data.type, title: ver.data.title, body: ver.data.body, attribution: ver.data.attribution || '', sessionNum: sNum, number: slNum, _courseName: verCourse, _sessionLabel: verSession };
        var newFrame = buildFrame(slide, x, y);
        parentNode.appendChild(newFrame);
        figma.currentPage.selection = [newFrame];
      }
      figma.ui.postMessage({ type: 'versionRestored', sessionNum: sNum, slideNum: slNum });
      sendSelectionData();
    } catch (err) {
      figma.ui.postMessage({ type: 'error', message: String(err) });
    }
  }
  if (msg.type === 'insertSlide') {
    try {
      await loadAllFonts();
      if (msg.colors) applyCustomColors(msg.colors);
      if (msg.fontConfig) await applyFontConfig(msg.fontConfig);
      var SPACING = W + 80;
      var ref = await figma.getNodeByIdAsync(msg.frameId);
      if (!ref) { figma.ui.postMessage({ type: 'error', message: 'Frame not found.' }); return; }
      var parentNode = ref.parent || figma.currentPage;

      // Determine insert position
      var insertX;
      if (msg.direction === 'left') {
        insertX = ref.x;
        // Shift selected frame and all frames to its right
        parentNode.children.forEach(function (child) {
          if (child.type === 'FRAME' && child.width === 1920 && child.x >= ref.x && child.name !== '[PREVIEW]') {
            child.x += SPACING;
          }
        });
      } else {
        insertX = ref.x + SPACING;
        // Shift all frames to the right of ref
        parentNode.children.forEach(function (child) {
          if (child.type === 'FRAME' && child.width === 1920 && child.x > ref.x && child.name !== '[PREVIEW]') {
            child.x += SPACING;
          }
        });
      }

      // Extract footer attributes from neighbor frame
      var neighborCourse = '', neighborSession = '';
      ref.findAll(function (n) { return n.type === 'TEXT'; }).forEach(function (t) {
        if (t.y > 950) {
          if (t.x < 960) neighborCourse = t.characters;
          else neighborSession = t.characters;
        }
      });

      // Build empty body slide with inherited footer
      var newSlide = {
        type: 'body',
        title: '',
        body: '',
        sessionNum: msg.sessionNum,
        number: 0,
        _courseName: neighborCourse,
        _sessionLabel: neighborSession
      };
      var newFrame = buildFrame(newSlide, insertX, ref.y);
      parentNode.appendChild(newFrame);
      figma.currentPage.selection = [newFrame];
      figma.viewport.scrollAndZoomIntoView([newFrame]);
      // Renumber page numbers after insert
      await renumberPageNumbers(parentNode);
      figma.ui.postMessage({ type: 'slideInserted', direction: msg.direction, slideNum: msg.slideNum });
    } catch (err) {
      figma.ui.postMessage({ type: 'error', message: String(err) });
    }
  }
  if (msg.type === 'deleteSlide') {
    try {
      var frame = await figma.getNodeByIdAsync(msg.frameId);
      if (!frame) { figma.ui.postMessage({ type: 'error', message: 'Frame not found.' }); return; }
      var parentNode = frame.parent || figma.currentPage;
      var deletedX = frame.x;
      var SPACING = W + 80;
      // Move to trash page (recoverable) before removing from canvas
      await moveToTrash(frame);
      frame.remove();
      // Remove preview too
      var prev = parentNode.findOne ? parentNode.findOne(function (n) { return n.name === '[PREVIEW]'; }) : null;
      if (prev) prev.remove();
      // Shift frames that were to the right of deleted slide
      parentNode.children.forEach(function (child) {
        if (child.type === 'FRAME' && child.width === 1920 && child.x > deletedX) {
          child.x -= SPACING;
        }
      });
      // Clear override if one existed
      if (msg.sessionNum && msg.slideNum) clearSlideOverride(msg.sessionNum, msg.slideNum);
      // Renumber page numbers after delete
      await renumberPageNumbers(parentNode);
      figma.ui.postMessage({ type: 'slideDeleted', trash: await getTrashList() });
    } catch (err) {
      figma.ui.postMessage({ type: 'error', message: String(err) });
    }
  }
  if (msg.type === 'previewSlide') {
    try {
      await loadAllFonts();
      if (msg.colors) applyCustomColors(msg.colors);
      if (msg.fontConfig) await applyFontConfig(msg.fontConfig);

      // Remove any existing preview frame
      var oldPreview = figma.currentPage.findOne(function (n) {
        return n.type === 'FRAME' && n.name === '[PREVIEW]';
      });
      if (oldPreview) oldPreview.remove();

      // Find original frame to extract page number
      var origPageNum = null;
      var origFrame = figma.currentPage.findOne(function (n) {
        return n.type === 'FRAME' && n.width === W && n.height === H && n.x === msg.frameX && Math.abs(n.y - msg.frameY) < 5 && n.name !== '[PREVIEW]';
      });
      if (origFrame) {
        var pgn = origFrame.findOne(function (n) { return n.name === '[PAGE_NUM]'; });
        if (pgn && pgn.type === 'TEXT') origPageNum = parseInt(pgn.characters);
      }

      // Build preview below the original frame
      var slide = {
        type: msg.slideType,
        title: msg.title,
        body: msg.body,
        attribution: msg.attribution || '',
        sessionNum: msg.sessionNum,
        number: msg.slideNum,
        _courseName: msg.courseName || '',
        _sessionLabel: msg.sessionLabel || '',
        _rawEdit: true,
        imageRef: msg.imageRef || '',
        _pageNum: origPageNum
      };
      var previewFrame = buildFrame(slide, msg.frameX, msg.frameY + 1080 + 60);
      previewFrame.name = '[PREVIEW]';
      previewFrame.opacity = 0.85;
      figma.currentPage.appendChild(previewFrame);
    } catch (err) {
      console.error('Preview error:', err);
      figma.ui.postMessage({ type: 'previewError', message: String(err) });
    }
  }
  if (msg.type === 'removePreview') {
    var prev = figma.currentPage.findOne(function (n) {
      return n.type === 'FRAME' && n.name === '[PREVIEW]';
    });
    if (prev) prev.remove();
  }
  if (msg.type === 'zoomToFrame') {
    var target = await figma.getNodeByIdAsync(msg.frameId);
    if (target) {
      figma.currentPage.selection = [target];
      figma.viewport.scrollAndZoomIntoView([target]);
    }
  }
  if (msg.type === 'getGithubToken') {
    var token = figma.root.getPluginData('githubToken') || '';
    figma.ui.postMessage({ type: 'githubToken', token: token });
  }
  if (msg.type === 'setGithubToken') {
    figma.root.setPluginData('githubToken', msg.token || '');
    figma.ui.postMessage({ type: 'githubToken', token: msg.token || '' });
  }
  if (msg.type === 'exportOverrides') {
    var allOverrides = getAllOverrides();
    figma.ui.postMessage({ type: 'overridesExport', courseId: msg.courseId, overrides: allOverrides });
  }
  if (msg.type === 'scanPage') {
    sendPageCharts();
  }
  if (msg.type === 'getFilteredOverrideCount') {
    var count = getFilteredOverrideCount(msg.sessionNums);
    figma.ui.postMessage({ type: 'overrideCount', count: count });
  }
  // V2: Notes
  if (msg.type === 'saveNote') {
    saveSlideNote(msg.sessionNum, msg.slideNum, msg.text);
    figma.ui.postMessage({ type: 'noteSaved', noteCount: getNoteCount() });
  }
  if (msg.type === 'deleteNote') {
    saveSlideNote(msg.sessionNum, msg.slideNum, '');
    figma.ui.postMessage({ type: 'noteDeleted', noteCount: getNoteCount() });
  }
  if (msg.type === 'getNote') {
    var noteData = getSlideNote(msg.sessionNum, msg.slideNum);
    figma.ui.postMessage({ type: 'noteData', sessionNum: msg.sessionNum, slideNum: msg.slideNum, note: noteData, noteCount: getNoteCount() });
  }
  if (msg.type === 'getAllNotes') {
    var allNotes = getAllPageNotes();
    figma.ui.postMessage({ type: 'allNotesData', notes: allNotes, noteCount: allNotes.length });
  }
  // V2: Review Flags
  if (msg.type === 'toggleReviewFlag') {
    var result = toggleReviewFlag(msg.sessionNum, msg.slideNum);
    figma.ui.postMessage({ type: 'reviewFlagToggled', sessionNum: msg.sessionNum, slideNum: msg.slideNum, flagged: result.flagged, totalFlags: result.totalFlags });
  }
  if (msg.type === 'getFlaggedSlides') {
    var flagged = getFlaggedSlides();
    figma.ui.postMessage({ type: 'flaggedSlidesList', slides: flagged, totalFlags: flagged.length });
  }
  // V2: Spell Check Ignores
  if (msg.type === 'saveSpellIgnores') {
    saveSpellIgnores(msg.sessionNum, msg.slideNum, msg.ignores);
    figma.ui.postMessage({ type: 'spellIgnoresSaved' });
  }
  if (msg.type === 'getSpellIgnores') {
    figma.ui.postMessage({ type: 'spellIgnoresData', sessionNum: msg.sessionNum, slideNum: msg.slideNum, ignores: getSpellIgnores(msg.sessionNum, msg.slideNum) });
  }
  // V2: Render source slide as image for Source View
  if (msg.type === 'renderSourceSlide') {
    try {
      await loadAllFonts();
      var srcSlide = {
        type: msg.slideType || 'body',
        title: msg.title || '',
        body: msg.body || '',
        sessionNum: msg.sessionNum,
        number: msg.slideNum,
        _courseName: msg.courseName || '',
        _sessionLabel: msg.sessionLabel || ''
      };
      // Build temp frame off-screen
      var tempFrame = buildFrame(srcSlide, -5000, -5000);
      figma.currentPage.appendChild(tempFrame);
      // Export as PNG at 0.5x for thumbnail
      var bytes = await tempFrame.exportAsync({ format: 'PNG', constraint: { type: 'SCALE', value: 0.4 } });
      tempFrame.remove();
      // Convert to base64
      var base64 = figma.base64Encode(bytes);
      figma.ui.postMessage({ type: 'sourceSlideImage', image: base64, sessionNum: msg.sessionNum, slideNum: msg.slideNum });
    } catch (err) {
      figma.ui.postMessage({ type: 'sourceSlideImage', image: null, error: String(err) });
    }
  }
  if (msg.type === 'getPageSlideList') {
    var slides = [];
    figma.currentPage.children.forEach(function (node) {
      if (node.type !== 'FRAME' || node.width !== 1920 || node.height !== 1080 || node.name === '[PREVIEW]') return;
      var nameMatch = node.name.match(/\[(\w+)\]\s*S(\d+)\s*\u00B7\s*(\d+)\s*\u2014\s*(.*)/);
      if (!nameMatch) return;
      var texts = [];
      node.findAll(function (n) { return n.type === 'TEXT'; }).forEach(function (t) {
        texts.push({ chars: t.characters, fontSize: t.fontSize, y: t.y });
      });
      texts.sort(function (a, b) { return a.y - b.y; });
      var title = '', body = '';
      for (var i = 0; i < texts.length; i++) {
        if (texts[i].y > 950) continue;
        if (texts[i].fontSize <= 32 && texts[i].y < 200 && !title) title = texts[i].chars;
        else if (texts[i].fontSize >= 32 && !body) body = texts[i].chars;
      }
      slides.push({ sessionNum: parseInt(nameMatch[2]), slideNum: parseInt(nameMatch[3]), title: title, body: body, frameId: node.id });
    });
    figma.ui.postMessage({ type: 'pageSlideList', slides: slides });
  }
  if (msg.type === 'saveSettings') {
    figma.clientStorage.setAsync('pluginSettings', msg.settings);
  }
  if (msg.type === 'loadSettings') {
    figma.clientStorage.getAsync('pluginSettings').then(function (settings) {
      figma.ui.postMessage({ type: 'savedSettings', settings: settings || null });
    });
  }
  if (msg.type === 'saveUnlock') {
    figma.clientStorage.setAsync('pluginUnlocked', true);
  }
  if (msg.type === 'checkUnlock') {
    figma.clientStorage.getAsync('pluginUnlocked').then(function (val) {
      figma.ui.postMessage({ type: 'unlockStatus', unlocked: !!val });
    });
  }
  if (msg.type === 'resize') {
    figma.ui.resize(480, Math.min(Math.max(msg.height, 480), 1200));
  }
  if (msg.type === 'getTrash') {
    figma.ui.postMessage({ type: 'trashList', items: await getTrashList() });
  }
  if (msg.type === 'recoverSlide') {
    try {
      var trashFrame = await figma.getNodeByIdAsync(msg.trashId);
      if (!trashFrame) { figma.ui.postMessage({ type: 'error', message: 'Trashed slide not found.' }); return; }
      var clone = trashFrame.clone();
      var origX = parseInt(trashFrame.getPluginData('_trash_x') || '0');
      var origY = parseInt(trashFrame.getPluginData('_trash_y') || '0');
      var SPACING = W + 80;

      // Shift any slide at or after the restore position to the right
      figma.currentPage.children.forEach(function (child) {
        if (child.type === 'FRAME' && child.width === W && child.height === H && child.x >= origX && child.name !== '[PREVIEW]') {
          child.x += SPACING;
        }
      });

      clone.x = origX;
      clone.y = origY;
      // Clear trash metadata
      clone.setPluginData('_trash_pageId', '');
      clone.setPluginData('_trash_pageName', '');
      clone.setPluginData('_trash_x', '');
      clone.setPluginData('_trash_y', '');
      clone.setPluginData('_trash_date', '');
      figma.currentPage.appendChild(clone);
      // Remove from trash
      trashFrame.remove();
      figma.currentPage.selection = [clone];
      figma.viewport.scrollAndZoomIntoView([clone]);
      figma.ui.postMessage({ type: 'slideRecovered', trash: await getTrashList() });
    } catch (err) {
      figma.ui.postMessage({ type: 'error', message: String(err) });
    }
  }
  if (msg.type === 'emptyTrash') {
    var pages = figma.root.children;
    for (var i = 0; i < pages.length; i++) {
      if (pages[i].name === '[TRASH]') {
        pages[i].remove();
        break;
      }
    }
    figma.ui.postMessage({ type: 'trashList', items: [] });
  }
  if (msg.type === 'renumberPages') {
    (async function () {
      await loadAllFonts();
      // Get all slide frames on current page sorted by x position
      var frames = figma.currentPage.children.filter(function (n) {
        return n.type === 'FRAME' && n.width === W && n.height === H && n.name !== '[PREVIEW]' && !/^\[COVER\]/.test(n.name);
      }).sort(function (a, b) { return a.x - b.x; });
      var count = 0;
      for (var i = 0; i < frames.length; i++) {
        var pgNode = frames[i].findOne(function (n) { return n.name === '[PAGE_NUM]'; });
        var pageNum = i + 2; // Page 1 is typically the cover
        if (pgNode && pgNode.type === 'TEXT') {
          await figma.loadFontAsync(pgNode.fontName);
          pgNode.characters = String(pageNum);
          count++;
        } else {
          // Add page number if it doesn't exist
          var pg = addText(frames[i], String(pageNum), {
            x: W - 100, y: H - 50, width: 80, height: 30,
            size: 18, color: COLORS.textFaint, opacity: 0.5,
            align: 'RIGHT', valign: 'BOTTOM'
          });
          if (pg) pg.name = '[PAGE_NUM]';
          count++;
        }
      }
      figma.ui.postMessage({ type: 'status', message: 'Renumbered ' + count + ' slides' });
    })();
  }
  if (msg.type === 'stripPageNumbers') {
    var frames = figma.currentPage.children.filter(function (n) {
      return n.type === 'FRAME' && n.width === W && n.height === H;
    });
    var count = 0;
    frames.forEach(function (f) {
      var pgNode = f.findOne(function (n) { return n.name === '[PAGE_NUM]'; });
      if (pgNode) { pgNode.remove(); count++; }
    });
    figma.ui.postMessage({ type: 'status', message: 'Stripped page numbers from ' + count + ' slides' });
  }
  if (msg.type === 'cleanupPrintPages') {
    var removed = cleanupPrintPages();
    figma.ui.postMessage({ type: 'printCleaned', count: removed });
  }
  if (msg.type === 'printCurrentPage') {
    try {
      await loadAllFonts();
      var srcPage = figma.currentPage;
      if (srcPage.name.indexOf('[PRINT]') === 0 || srcPage.name.indexOf('[TRASH]') === 0 || srcPage.name.indexOf('[STORAGE]') === 0) {
        figma.ui.postMessage({ type: 'error', message: 'Switch to a slide page first' });
      } else {
        var printPage = await createPrintPage(srcPage);
        figma.currentPage = printPage;
        figma.ui.postMessage({ type: 'printPagesReady' });
      }
    } catch (err) {
      figma.ui.postMessage({ type: 'error', message: String(err) });
    }
  }
  if (msg.type === 'cancel') {
    figma.closePlugin();
  }
};

// ============================================================
// COLORS & LAYOUT
// ============================================================

var COLORS = {
  bg:          { r: 0.482, g: 0.463, b: 0.412 }, // #7B7669
  textPrimary: { r: 0.906, g: 0.890, b: 0.859 }, // #E7E3DB
  textBody:    { r: 0.906, g: 0.890, b: 0.859 }, // #E7E3DB
  textMuted:   { r: 0.906, g: 0.890, b: 0.859 }, // #E7E3DB (used at 50% opacity for refs)
  textFaint:   { r: 0.906, g: 0.890, b: 0.859 }, // #E7E3DB (used at 50% opacity for footer)
  redMarker:   { r: 0.886, g: 0.263, b: 0.216 }, // #E24337 — illustration needed
};

function hexToRgb(hex) {
  hex = hex.replace(/^#/, '');
  return {
    r: parseInt(hex.substring(0, 2), 16) / 255,
    g: parseInt(hex.substring(2, 4), 16) / 255,
    b: parseInt(hex.substring(4, 6), 16) / 255
  };
}

function applyCustomColors(colors) {
  if (colors.bg) COLORS.bg = hexToRgb(colors.bg);
  if (colors.text) {
    COLORS.textPrimary = hexToRgb(colors.text);
    COLORS.textBody = hexToRgb(colors.text);
  }
  if (colors.footer) COLORS.textFaint = hexToRgb(colors.footer);
  if (colors.attribution) COLORS.textMuted = hexToRgb(colors.attribution);
  if (colors.marker) COLORS.redMarker = hexToRgb(colors.marker);
}

// ============================================================
// SLIDE OVERRIDES — persist edits in the Figma file
// Two modes:
//   'text'   — title/body/type stored as JSON, re-rendered via buildFrame
//   'design' — entire frame cloned to hidden storage page, copied on rebuild
// ============================================================

var OVERRIDE_PREFIX = 'slideOverride_';
var STORAGE_PAGE_NAME = '_Slide Overrides (do not delete)';

function overrideKey(sessionNum, slideNum) {
  return OVERRIDE_PREFIX + 'S' + sessionNum + '_' + slideNum;
}

async function getStoragePage() {
  for (var i = 0; i < figma.root.children.length; i++) {
    if (figma.root.children[i].name === STORAGE_PAGE_NAME) {
      await figma.root.children[i].loadAsync();
      return figma.root.children[i];
    }
  }
  var page = figma.createPage();
  page.name = STORAGE_PAGE_NAME;
  return page;
}

async function saveSlideOverride(sessionNum, slideNum, data) {
  var key = overrideKey(sessionNum, slideNum);
  var existing = getSlideOverride(sessionNum, slideNum);

  // Push current state to version history before overwriting
  var vKey = key + '_versions';
  var versions = JSON.parse(figma.root.getPluginData(vKey) || '[]');
  if (existing) {
    var now = new Date();
    var ts = (now.getMonth()+1) + '/' + now.getDate() + ' ' + now.getHours() + ':' + String(now.getMinutes()).padStart(2,'0');
    var versionLabel = 'v' + (versions.length + 1) + ' \u2014 ' + (existing.mode === 'design' ? 'Design' : 'Text') + ' \u00B7 ' + ts;
    versions.push({ label: versionLabel, data: existing, timestamp: now.toISOString() });
    // Keep max 10 versions
    if (versions.length > 10) versions = versions.slice(versions.length - 10);
    figma.root.setPluginData(vKey, JSON.stringify(versions));
  }

  // If replacing a design override, remove the old stored frame
  if (existing && existing.mode === 'design' && existing.committedNodeId) {
    var old = await figma.getNodeByIdAsync(existing.committedNodeId);
    if (old) old.remove();
  }
  figma.root.setPluginData(key, JSON.stringify(data));
  var index = JSON.parse(figma.root.getPluginData('overrideIndex') || '[]');
  if (index.indexOf(key) === -1) index.push(key);
  figma.root.setPluginData('overrideIndex', JSON.stringify(index));
}

function getSlideVersions(sessionNum, slideNum) {
  var key = overrideKey(sessionNum, slideNum) + '_versions';
  try { return JSON.parse(figma.root.getPluginData(key) || '[]'); } catch (e) { return []; }
}

async function commitDesignOverride(sessionNum, slideNum, frame) {
  var storagePage = await getStoragePage();
  var clone = frame.clone();
  clone.name = '[COMMITTED] S' + sessionNum + ' \u00B7 ' + slideNum;
  storagePage.appendChild(clone);

  // Strip "Illustration Needed" marker from committed chart/graphic slides
  var isChart = /^\[(CHART|GRAPHIC)\]/i.test(frame.name);
  if (isChart) {
    var toRemove = [];
    clone.findAll(function (n) {
      if (n.type === 'TEXT' && n.characters === 'Illustration Needed') toRemove.push(n);
      if (n.type === 'ELLIPSE' && n.width === 24 && n.height === 24) toRemove.push(n);
      return false;
    });
    toRemove.forEach(function (n) { n.remove(); });
  }
  // Stack committed frames vertically on storage page
  var yOffset = 0;
  for (var i = 0; i < storagePage.children.length; i++) {
    var child = storagePage.children[i];
    if (child.id !== clone.id) {
      var bottom = child.y + child.height + 40;
      if (bottom > yOffset) yOffset = bottom;
    }
  }
  clone.x = 0;
  clone.y = yOffset;

  saveSlideOverride(sessionNum, slideNum, {
    mode: 'design',
    committedNodeId: clone.id
  });
  return clone.id;
}

function getSlideOverride(sessionNum, slideNum) {
  var key = overrideKey(sessionNum, slideNum);
  var raw = figma.root.getPluginData(key);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

async function clearSlideOverride(sessionNum, slideNum) {
  var existing = getSlideOverride(sessionNum, slideNum);
  // Clean up stored frame if design mode
  if (existing && existing.mode === 'design' && existing.committedNodeId) {
    var stored = await figma.getNodeByIdAsync(existing.committedNodeId);
    if (stored) stored.remove();
  }
  // Also clean up design frames in version history
  var vKey = overrideKey(sessionNum, slideNum) + '_versions';
  var versions = JSON.parse(figma.root.getPluginData(vKey) || '[]');
  for (var vi = 0; vi < versions.length; vi++) {
    var vd = versions[vi].data;
    if (vd && vd.mode === 'design' && vd.committedNodeId) {
      var vNode = await figma.getNodeByIdAsync(vd.committedNodeId);
      if (vNode) vNode.remove();
    }
  }
  figma.root.setPluginData(vKey, '');

  var key = overrideKey(sessionNum, slideNum);
  figma.root.setPluginData(key, '');
  var index = JSON.parse(figma.root.getPluginData('overrideIndex') || '[]');
  index = index.filter(function (k) { return k !== key; });
  figma.root.setPluginData('overrideIndex', JSON.stringify(index));
}

async function clearAllOverrides() {
  var index = JSON.parse(figma.root.getPluginData('overrideIndex') || '[]');
  for (var i = 0; i < index.length; i++) {
    var raw = figma.root.getPluginData(index[i]);
    if (raw) {
      try {
        var data = JSON.parse(raw);
        if (data.mode === 'design' && data.committedNodeId) {
          var stored = await figma.getNodeByIdAsync(data.committedNodeId);
          if (stored) stored.remove();
        }
      } catch (e) {}
    }
    figma.root.setPluginData(index[i], '');
  }
  figma.root.setPluginData('overrideIndex', '[]');
  // Remove storage page if empty
  var storagePage = null;
  for (var j = 0; j < figma.root.children.length; j++) {
    if (figma.root.children[j].name === STORAGE_PAGE_NAME) {
      storagePage = figma.root.children[j];
      break;
    }
  }
  if (storagePage && storagePage.children.length === 0) storagePage.remove();
}

function getOverrideCount() {
  return JSON.parse(figma.root.getPluginData('overrideIndex') || '[]').length;
}

function getFilteredOverrideCount(sessionNums) {
  if (!sessionNums || sessionNums.length === 0) return getOverrideCount();
  var index = JSON.parse(figma.root.getPluginData('overrideIndex') || '[]');
  var count = 0;
  for (var i = 0; i < index.length; i++) {
    var match = index[i].match(/S(\d+)_(\d+)$/);
    if (match && sessionNums.indexOf(parseInt(match[1])) !== -1) count++;
  }
  return count;
}

function getAllOverrides() {
  var index = JSON.parse(figma.root.getPluginData('overrideIndex') || '[]');
  var result = [];
  for (var i = 0; i < index.length; i++) {
    var raw = figma.root.getPluginData(index[i]);
    if (raw) {
      try {
        var data = JSON.parse(raw);
        var match = index[i].match(/S(\d+)_(\d+)$/);
        if (match) {
          data.sessionNum = parseInt(match[1]);
          data.slideNum = parseInt(match[2]);
        }
        result.push(data);
      } catch (e) {}
    }
  }
  return result;
}

// ============================================================
// V2: SLIDE NOTES
// ============================================================

function noteKey(sNum, slNum) { return 'slideNote_S' + sNum + '_' + slNum; }

function saveSlideNote(sNum, slNum, text) {
  var key = noteKey(sNum, slNum);
  if (!text || !text.trim()) {
    figma.root.setPluginData(key, '');
    var index = JSON.parse(figma.root.getPluginData('noteIndex') || '[]');
    index = index.filter(function (k) { return k !== key; });
    figma.root.setPluginData('noteIndex', JSON.stringify(index));
    return;
  }
  var userName = 'Unknown';
  try { userName = figma.currentUser ? figma.currentUser.name : 'Unknown'; } catch (e) {}
  var noteData = JSON.stringify({
    text: text,
    author: userName,
    timestamp: new Date().toISOString()
  });
  figma.root.setPluginData(key, noteData);
  var index = JSON.parse(figma.root.getPluginData('noteIndex') || '[]');
  if (index.indexOf(key) === -1) index.push(key);
  figma.root.setPluginData('noteIndex', JSON.stringify(index));
}

function getSlideNote(sNum, slNum) {
  var raw = figma.root.getPluginData(noteKey(sNum, slNum)) || '';
  if (!raw) return null;
  try {
    var parsed = JSON.parse(raw);
    if (parsed.text) return parsed; // new format {text, author, timestamp}
    return null;
  } catch (e) {
    // Legacy plain text note
    return raw.trim() ? { text: raw, author: 'Unknown', timestamp: '' } : null;
  }
}

function getNoteCount() {
  var index = JSON.parse(figma.root.getPluginData('noteIndex') || '[]');
  return index.filter(function (k) {
    var raw = figma.root.getPluginData(k) || '';
    if (!raw.trim()) return false;
    try { var p = JSON.parse(raw); return p.text && p.text.trim(); } catch (e) { return raw.trim() !== ''; }
  }).length;
}

function getAllPageNotes() {
  // Find all slides on current page and return their notes
  var results = [];
  var index = JSON.parse(figma.root.getPluginData('noteIndex') || '[]');
  index.forEach(function (key) {
    var match = key.match(/slideNote_S(\d+)_(\d+)$/);
    if (!match) return;
    var sNum = parseInt(match[1]);
    var slNum = parseInt(match[2]);
    var note = getSlideNote(sNum, slNum);
    if (!note) return;
    // Find frame on current page
    var nameSnippet = 'S' + sNum + ' \u00B7 ' + slNum + ' \u2014';
    var frameId = null, frameName = '';
    figma.currentPage.children.forEach(function (node) {
      if (node.type === 'FRAME' && node.name.indexOf(nameSnippet) !== -1 && node.name !== '[PREVIEW]') {
        frameId = node.id;
        frameName = node.name;
      }
    });
    if (frameId) {
      results.push({ sessionNum: sNum, slideNum: slNum, frameId: frameId, frameName: frameName, note: note });
    }
  });
  return results;
}

// ============================================================
// V2: REVIEW FLAGS
// ============================================================

function getReviewFlags() {
  try { return JSON.parse(figma.root.getPluginData('reviewFlags') || '[]'); } catch (e) { return []; }
}

function isSlideReviewFlagged(sNum, slNum) {
  return getReviewFlags().some(function (f) { return f.sessionNum === sNum && f.slideNum === slNum; });
}

function toggleReviewFlag(sNum, slNum) {
  var flags = getReviewFlags();
  var idx = -1;
  for (var i = 0; i < flags.length; i++) {
    if (flags[i].sessionNum === sNum && flags[i].slideNum === slNum) { idx = i; break; }
  }
  var nowFlagged;
  if (idx !== -1) { flags.splice(idx, 1); nowFlagged = false; }
  else { flags.push({ sessionNum: sNum, slideNum: slNum }); nowFlagged = true; }
  figma.root.setPluginData('reviewFlags', JSON.stringify(flags));

  // Update frame name on canvas
  var nameSnippet = 'S' + sNum + ' \u00B7 ' + slNum + ' \u2014';
  figma.currentPage.children.forEach(function (node) {
    if (node.type === 'FRAME' && node.name.indexOf(nameSnippet) !== -1 && node.name !== '[PREVIEW]') {
      if (nowFlagged && node.name.indexOf('\u2691') === -1) {
        node.name = '\u2691 ' + node.name;
      } else if (!nowFlagged) {
        node.name = node.name.replace(/^\u2691\s*/, '');
      }
    }
  });
  return { flagged: nowFlagged, totalFlags: flags.length };
}

function getFlaggedSlides() {
  var flags = getReviewFlags();
  var results = [];
  flags.forEach(function (f) {
    var nameSnippet = 'S' + f.sessionNum + ' \u00B7 ' + f.slideNum + ' \u2014';
    figma.currentPage.children.forEach(function (node) {
      if (node.type === 'FRAME' && node.name.indexOf(nameSnippet) !== -1 && node.name !== '[PREVIEW]') {
        results.push({ sessionNum: f.sessionNum, slideNum: f.slideNum, frameId: node.id, name: node.name });
      }
    });
  });
  return results;
}

// ============================================================
// V2: SPELL CHECK IGNORES
// ============================================================

function spellIgnoreKey(sNum, slNum) { return 'spellIgnores_S' + sNum + '_' + slNum; }

function saveSpellIgnores(sNum, slNum, ignores) {
  figma.root.setPluginData(spellIgnoreKey(sNum, slNum), JSON.stringify(ignores || []));
}

function getSpellIgnores(sNum, slNum) {
  try { return JSON.parse(figma.root.getPluginData(spellIgnoreKey(sNum, slNum)) || '[]'); } catch (e) { return []; }
}

// ============================================================
// TRASH (recoverable deleted slides)
// ============================================================

async function getTrashPage() {
  var pages = figma.root.children;
  for (var i = 0; i < pages.length; i++) {
    if (pages[i].name === '[TRASH]') {
      await pages[i].loadAsync();
      return pages[i];
    }
  }
  var page = figma.createPage();
  page.name = '[TRASH]';
  return page;
}

async function moveToTrash(frame) {
  var trashPage = await getTrashPage();
  var clone = frame.clone();
  // Store original location metadata
  clone.setPluginData('_trash_pageId', frame.parent.id || '');
  clone.setPluginData('_trash_pageName', (frame.parent && frame.parent.name) || '');
  clone.setPluginData('_trash_x', String(frame.x));
  clone.setPluginData('_trash_y', String(frame.y));
  clone.setPluginData('_trash_date', new Date().toISOString().slice(0, 16).replace('T', ' '));
  // Stack trashed frames vertically
  var trashY = 0;
  trashPage.children.forEach(function (c) {
    if (c.type === 'FRAME') {
      var bottom = c.y + c.height + 40;
      if (bottom > trashY) trashY = bottom;
    }
  });
  clone.x = 0;
  clone.y = trashY;
  trashPage.appendChild(clone);
  return clone;
}

async function getTrashList() {
  var pages = figma.root.children;
  var trashPage = null;
  for (var i = 0; i < pages.length; i++) {
    if (pages[i].name === '[TRASH]') { trashPage = pages[i]; break; }
  }
  if (!trashPage) return [];
  await trashPage.loadAsync();
  var items = [];
  trashPage.children.forEach(function (c) {
    if (c.type === 'FRAME') {
      items.push({
        id: c.id,
        name: c.name,
        pageName: c.getPluginData('_trash_pageName') || '',
        date: c.getPluginData('_trash_date') || ''
      });
    }
  });
  return items;
}

// Renumber page numbers on all slide frames on the current page.
// Cover (first frame by X) = no number, rest numbered starting at 2.
async function renumberPageNumbers(parentNode) {
  var slideFrames = [];
  parentNode.children.forEach(function (child) {
    if (child.type === 'FRAME' && child.width === 1920 && child.height === 1080 && child.name !== '[PREVIEW]') {
      slideFrames.push(child);
    }
  });
  slideFrames.sort(function (a, b) { return a.x - b.x; });
  for (var i = 0; i < slideFrames.length; i++) {
    var frame = slideFrames[i];
    var pageNumNode = frame.findOne(function (n) { return n.name === '[PAGE_NUM]'; });
    if (i === 0) {
      // Cover slide — remove page number if present
      if (pageNumNode) pageNumNode.remove();
    } else {
      if (pageNumNode) {
        await figma.loadFontAsync(pageNumNode.fontName);
        pageNumNode.characters = String(i + 1);
      }
    }
  }
}

// Build or clone a slide frame, respecting overrides
// Returns the frame to append to the page
async function buildOrCloneFrame(slide, x, y) {
  var override = getSlideOverride(slide.sessionNum, slide.number);
  if (override && override.mode === 'design' && override.committedNodeId) {
    var stored = await figma.getNodeByIdAsync(override.committedNodeId);
    if (stored) {
      var clone = stored.clone();
      clone.x = x;
      clone.y = y;
      // Update page number on cloned design override if needed
      if (slide._pageNum) {
        var pgNode = clone.findOne(function (n) { return n.name === '[PAGE_NUM]'; });
        if (pgNode) {
          pgNode.characters = String(slide._pageNum);
        }
        // If no page number node exists, it will be added by the caller if needed
      }
      return clone;
    }
  }
  // Text override or no override — apply text overrides and build normally
  if (override && override.mode === 'text') {
    if (override.type) slide.type = override.type;
    if (typeof override.title === 'string') slide.title = override.title;
    if (typeof override.body === 'string') slide.body = override.body;
  }
  return buildFrame(slide, x, y);
}

// Detect scripture references — used to prevent them from appearing as slide titles
function isScriptureTitle(s) {
  if (!s) return false;
  return /^[123]?\s*(Gen|Exod|Lev|Num|Deut|Josh|Judg|Ruth|Sam|Kgs|Chr|Ezra|Neh|Esth|Job|Ps|Prov|Eccl|Song|Isa|Jer|Lam|Ezek|Dan|Hos|Joel|Amos|Obad|Jon|Mic|Nah|Hab|Zeph|Hag|Zech|Mal|Matt|Mark|Luke|John|Acts|Rom|Cor|Gal|Eph|Phil|Col|Thess|Tim|Titus|Phlm|Heb|Jas|Pet|Jude|Rev|Genesis|Exodus|Leviticus|Numbers|Deuteronomy|Joshua|Judges|Ruth|Samuel|Kings|Chronicles|Ezra|Nehemiah|Esther|Job|Psalms?|Proverbs|Ecclesiastes|Isaiah|Jeremiah|Lamentations|Ezekiel|Daniel|Hosea|Joel|Amos|Obadiah|Jonah|Micah|Nahum|Habakkuk|Zephaniah|Haggai|Zechariah|Malachi|Matthew|Mark|Luke|John|Acts|Romans|Corinthians|Galatians|Ephesians|Philippians|Colossians|Thessalonians|Timothy|Titus|Philemon|Hebrews|James|Peter|Jude|Revelation)\s*\.?\s*\d/i.test(s.trim());
}

var W = 1920;
var H = 1080;

var CONTENT_W = 1400;
var SIDE_MARGIN = (W - CONTENT_W) / 2;  // 260px — centers 1400px content
var FOOTER_Y = H - 74;                  // 50px from bottom edge for 24px text
var FOOTER_MARGIN = 50;                 // footer inset from edges
var LABEL_Y = 50;
var CONTENT_TOP = 120;
var CONTENT_BOTTOM = H - 100;
var CONTENT_H = CONTENT_BOTTOM - CONTENT_TOP;

// Text sizes
var TITLE_SIZE = 32;
var FOOTER_SIZE = 24;
var REF_OFFSET = 8;       // attribution/ref is this many px smaller than body
var REF_GAP = 64;         // gap between body text bottom and attribution

// Body text scales down in 4-8px steps by content length
// Prioritize legibility — never go below 32px
var SCALE_THRESHOLDS = [
  { max: 150, size: 48 },
  { max: 300, size: 44 },
  { max: 500, size: 40 },
  { max: 750, size: 36 },
  { max: 9999, size: 32 }
];

function bodySize(text) {
  var len = text.length;
  for (var i = 0; i < SCALE_THRESHOLDS.length; i++) {
    if (len <= SCALE_THRESHOLDS[i].max) return SCALE_THRESHOLDS[i].size;
  }
  return 24;
}

// ============================================================
// FONTS — try Signifier + Untitled Sans, fall back to Inter
// ============================================================

var FONTS = {
  serifRegular:      { family: 'Signifier', style: 'Regular' },
  serifMedium:       { family: 'Signifier', style: 'Medium' },
  serifItalic:       { family: 'Signifier', style: 'RegularItalic' },
  serifMediumItalic: { family: 'Signifier', style: 'MediumItalic' },
  serifBold:         { family: 'Signifier', style: 'Medium' },  // User prefers Medium over Bold
  sansRegular:       { family: 'Untitled Sans', style: 'Regular' },
  sansMedium:        { family: 'Untitled Sans', style: 'Medium' },
  sansBold:          { family: 'Untitled Sans', style: 'Bold' },
};

// Secondary attempt (some installs name italic differently)
var FONTS_ALT = {
  serifItalic:       { family: 'Signifier', style: 'Regular Italic' },
  serifMediumItalic: { family: 'Signifier', style: 'Medium Italic' },
};

var FALLBACKS = {
  serifRegular:      { family: 'Inter', style: 'Regular' },
  serifMedium:       { family: 'Inter', style: 'Medium' },
  serifItalic:       { family: 'Inter', style: 'Italic' },
  serifMediumItalic: { family: 'Inter', style: 'Medium Italic' },
  serifBold:         { family: 'Inter', style: 'Bold' },
  sansRegular:       { family: 'Inter', style: 'Regular' },
  sansMedium:        { family: 'Inter', style: 'Medium' },
  sansBold:          { family: 'Inter', style: 'Bold' },
};

var RESOLVED_FONTS = {};

async function loadAllFonts() {
  var keys = Object.keys(FONTS);
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    try {
      await figma.loadFontAsync(FONTS[key]);
      RESOLVED_FONTS[key] = FONTS[key];
    } catch (e) {
      // Try alternate name if available
      if (FONTS_ALT[key]) {
        try {
          await figma.loadFontAsync(FONTS_ALT[key]);
          RESOLVED_FONTS[key] = FONTS_ALT[key];
          continue;
        } catch (e2) { /* fall through to Inter */ }
      }
      await figma.loadFontAsync(FALLBACKS[key]);
      RESOLVED_FONTS[key] = FALLBACKS[key];
    }
  }
}

// Font role overrides — set from UI font selectors
// Maps role names to RESOLVED_FONTS keys or direct FontName objects
var FONT_ROLES = {
  title: null,       // slide titles — defaults to serifMedium
  body: null,        // body/quote text — defaults to serifRegular
  emphasis: null,    // *emphasis* markers — defaults to serifMediumItalic
  attribution: null, // scripture refs, quote attribution — defaults to serifItalic
  labels: null       // footer, cover class labels — defaults to sansRegular
};

async function applyFontConfig(config) {
  var roleMap = {
    title: 'fontTitle',
    body: 'fontBody',
    emphasis: 'fontEmphasis',
    attribution: 'fontAttribution',
    labels: 'fontLabels'
  };
  for (var role in config) {
    if (config[role] && config[role].family) {
      var font = { family: config[role].family, style: config[role].style };
      try {
        await figma.loadFontAsync(font);
        FONT_ROLES[role] = font;
      } catch (e) {
        // Try alternate style names (e.g., "MediumItalic" → "Medium Italic")
        var altStyle = config[role].style.replace(/([a-z])([A-Z])/g, '$1 $2');
        if (altStyle !== config[role].style) {
          try {
            var altFont = { family: config[role].family, style: altStyle };
            await figma.loadFontAsync(altFont);
            FONT_ROLES[role] = altFont;
          } catch (e2) { /* keep default */ }
        }
      }
    }
  }
}

// ============================================================
// TEXT POLISHING — proper punctuation, spacing, clean dashes
// ============================================================

function polishText(text) {
  var s = text;

  // Unwrap PDF soft-wrap line breaks: single \n → space
  // Only keep a break when the previous line ends with terminal punctuation
  // and the next line starts a new thought (list item or uppercase start).
  // Everything else is treated as a PDF soft-wrap to merge.
  s = s.replace(/\n{2,}/g, '\u0000PARA\u0000');
  var lines = s.split('\n');
  var merged = [lines[0]];
  for (var i = 1; i < lines.length; i++) {
    var line = lines[i];
    var trimmed = line.trim();
    var prevLine = merged[merged.length - 1].trim();
    var prevEndsTerminal = /[.;!?)\u201D"']$/.test(prevLine);
    var isList = /^(\d+\.|[●•○◦\-\*]\s)/.test(trimmed);
    var startsQuote = /^[\u201C\u2018"]/.test(trimmed);

    if (isList || startsQuote || (prevEndsTerminal && /^[A-Z\d]/.test(trimmed))) {
      // New point: previous sentence ended AND this line starts fresh
      merged.push('\n' + trimmed);
    } else {
      // Continuation — merge (PDF soft-wrap)
      merged.push(' ' + trimmed);
    }
  }
  s = merged.join('');
  s = s.replace(/\u0000PARA\u0000/g, '\n\n');

  // Clean up bullet characters — remove heavy bullets, keep text only
  // Mark sub-items (○) with indent before stripping
  s = s.replace(/[○◦]\s*/g, '  SUB_ITEM ');
  s = s.replace(/[●•]\s*/g, '');

  // Fix arrows: -> to ›, => to ›, --> to ›
  s = s.replace(/\s*-{1,2}>\s*/g, ' \u203A ');
  s = s.replace(/\s*=>\s*/g, ' \u203A ');

  // Normalize dashes: replace -- with em dash
  s = s.replace(/\s*--\s*/g, '\u2014');

  // Ensure proper em dash spacing (thin space or no space)
  s = s.replace(/\s*\u2014\s*/g, '\u2014');

  // Straight quotes to curly
  s = s.replace(/"([^"]*?)"/g, '\u201C$1\u201D');
  s = s.replace(/'([^']*?)'/g, '\u2018$1\u2019');

  // Strip inline verse numbers (superscripts pasted as plain digits)
  // e.g. "garden; 17 but you must" → "garden; but you must"
  // Matches digits after punctuation+space or at start of quoted text
  s = s.replace(/([;,.!?]\s*)\d{1,3}\s+(?=[a-z])/g, '$1');
  s = s.replace(/([\u201C\u2018"'])\s*\d{1,3}\s+/g, '$1');

  // Strip editorial notes like (italics added), (emphasis added), etc.
  s = s.replace(/\s*\(italics\s+added\)\s*/gi, ' ');
  s = s.replace(/\s*\(emphasis\s+(added|mine)\)\s*/gi, ' ');

  // Fix double spaces
  s = s.replace(/  +/g, ' ');

  // Fix common typos from PDF extraction
  s = s.replace(/\bWorhsip\b/g, 'Worship');
  s = s.replace(/\bdistincy\b/g, 'distinct');
  s = s.replace(/\bTrintiy\b/g, 'Trinity');

  // Expand scripture book abbreviations to full names
  s = expandBookNames(s);

  // AP style: no double periods
  s = s.replace(/\.{2,}/g, '.');

  // Orphan prevention: join last 3 words of each paragraph/line with
  // non-breaking spaces so they can't be split across lines
  s = s.replace(/[^\n]+/g, function (line) {
    var words = line.split(/\s+/);
    if (words.length >= 4) {
      var last3 = words.splice(-3).join('\u00A0');
      return words.join(' ') + ' ' + last3;
    }
    // For very short lines (3 words or fewer), join all with nbsp
    if (words.length > 1) {
      return words.join('\u00A0');
    }
    return line;
  });

  return s.trim();
}

// Ensure text ends with proper punctuation (AP style)
function ensurePeriod(text) {
  var s = text.trim();
  if (s.length === 0) return s;
  var lastChar = s[s.length - 1];
  // Don't add period if already ends with punctuation or closing paren
  if (/[.!?:;\u201D"'\u2019)]$/.test(s)) return s;
  return s + '.';
}

function cleanAttribution(attr) {
  // Strip all leading dashes, hyphens, em/en dashes
  var s = attr.replace(/^[\s\-—–]+/, '').trim();
  // Expand any abbreviated book names in attributions
  s = expandBookNames(s);
  return s;
}

// ============================================================
// PRINT VIEW — white bg, black text, export as PDF
// ============================================================

async function createPrintPage(sourcePage) {
  var printPage = figma.createPage();
  printPage.name = '[PRINT] ' + sourcePage.name;
  printPage.backgrounds = [{ type: 'SOLID', color: { r: 0.961, g: 0.961, b: 0.961 } }];

  var slideFrames = [];
  sourcePage.children.forEach(function (child) {
    if (child.type === 'FRAME' && child.width === 1920 && child.height === 1080) {
      slideFrames.push(child);
    }
  });
  slideFrames.sort(function (a, b) { return a.x - b.x; });

  for (var i = 0; i < slideFrames.length; i++) {
    var clone = slideFrames[i].clone();
    clone.x = i * (1920 + 40);
    clone.y = 0;
    // White slide background
    clone.fills = [{ type: 'SOLID', color: { r: 1, g: 1, b: 1 } }];
    // Convert all text to black, preserving original opacity
    clone.findAll(function (n) { return n.type === 'TEXT'; }).forEach(function (t) {
      t.fills = [{ type: 'SOLID', color: { r: 0, g: 0, b: 0 } }];
      // Keep original opacity — don't force to 1
    });
    // Ensure page number on print slides (starting from page 2)
    if (i > 0) {
      var pgNode = clone.findOne(function (n) { return n.name === '[PAGE_NUM]'; });
      if (!pgNode) {
        // Add page number if not already present
        await loadAllFonts();
        var pgText = addText(clone, String(i + 1), {
          x: (W - 100) / 2,
          y: FOOTER_Y,
          w: 100,
          h: 32,
          size: FOOTER_SIZE,
          color: { r: 0, g: 0, b: 0 },
          opacity: 1,
          align: 'CENTER',
          font: 'sans'
        });
        if (pgText) {
          pgText.name = '[PAGE_NUM]';
          pgText.fills = [{ type: 'SOLID', color: { r: 0, g: 0, b: 0 } }];
        }
      }
    }
    printPage.appendChild(clone);
  }
  return printPage;
}

// No auto-export — print pages are kept for user to export via File > Export Frames to PDF

function cleanupPrintPages() {
  var removed = 0;
  var pages = figma.root.children.slice();
  // Switch away from print page before deleting
  var isPrintCurrent = figma.currentPage.name.indexOf('[PRINT]') === 0;
  if (isPrintCurrent) {
    for (var j = 0; j < pages.length; j++) {
      if (pages[j].name.indexOf('[PRINT]') !== 0 && pages[j].name.indexOf('[TRASH]') !== 0 && pages[j].name.indexOf('[STORAGE]') !== 0) {
        figma.currentPage = pages[j];
        break;
      }
    }
  }
  for (var i = 0; i < pages.length; i++) {
    if (pages[i].name.indexOf('[PRINT]') === 0) {
      pages[i].remove();
      removed++;
    }
  }
  return removed;
}

// ============================================================
// MAIN BUILD
// ============================================================

async function buildSlides(slides, sessionNum, sessionName, courseName, logoData, totalSessions, includePageNumbers) {
  await loadAllFonts();

  // Pre-compute consistent font sizes for progressive slides (same title).
  // The largest body (last slide in sequence) determines the size for all.
  var titleGroupMaxLen = {};
  for (var g = 0; g < slides.length; g++) {
    var sl = slides[g];
    if ((sl.type === 'body' || sl.type === 'list') && sl.title) {
      var key = sl.title;
      var len = (sl.body || '').length;
      if (!titleGroupMaxLen[key] || len > titleGroupMaxLen[key]) {
        titleGroupMaxLen[key] = len;
      }
    }
  }
  // Store the unified size on each slide in the group
  for (var g2 = 0; g2 < slides.length; g2++) {
    var sl2 = slides[g2];
    if ((sl2.type === 'body' || sl2.type === 'list') && sl2.title && titleGroupMaxLen[sl2.title]) {
      sl2._groupSize = bodySize({ length: titleGroupMaxLen[sl2.title] });
    }
  }

  // Always create a new page — never touch the current page
  var page = figma.createPage();

  // Set page background to light gray
  page.backgrounds = [{ type: 'SOLID', color: { r: 0.961, g: 0.961, b: 0.961 } }]; // #F5F5F5

  // Rename the page to the course + class
  var pageLabel = courseName || '';
  if (sessionName) pageLabel += ' \u2014 Class ' + sessionNum + ': ' + cleanSessionLabel(sessionName);
  if (pageLabel) page.name = pageLabel;

  var builtCount = 0;

  // Insert cover slide at the start of the session
  var coverFrame = buildCoverSlide(courseName, sessionNum, sessionName, 0, 0, logoData, totalSessions);
  page.appendChild(coverFrame);
  builtCount++;

  for (var index = 0; index < slides.length; index++) {
    var slide = slides[index];
    slide._courseName = courseName;
    slide._sessionLabel = sessionName;
    // Page numbers: cover is page 1 (no number), content slides start at page 2
    if (includePageNumbers) slide._pageNum = index + 2;
    var x = (index + 1) * (W + 80);
    var frame = buildOrCloneFrame(slide, x, 0);
    page.appendChild(frame);
    builtCount++;
  }

  // Switch to the new page and zoom to fit
  figma.currentPage = page;
  figma.viewport.scrollAndZoomIntoView(page.children);

  figma.ui.postMessage({
    type: 'done',
    count: builtCount,
    session: sessionNum
  });
  sendPageCharts();
}

// ============================================================
// BUILD ALL — one page per session
// ============================================================

async function buildAllSessions(sessionGroups, courseName, logoData, totalSessions, includePageNumbers) {
  console.log('[PLUGIN] buildAllSessions called with', sessionGroups.length, 'groups');
  await loadAllFonts();

  var totalSlides = 0;
  var numSessions = sessionGroups.length;
  if (!totalSessions) totalSessions = numSessions;

  var errors = [];
  for (var i = 0; i < sessionGroups.length; i++) {
    var group = sessionGroups[i];
    var slides = group.slides;
    var sessionNum = group.sessionNum;
    var sessionName = group.sessionName;

    figma.ui.postMessage({ type: 'progress', current: i + 1, total: totalSessions });

    try {
      // Always create a new page — never touch the current page
      var page = figma.createPage();

      // Set page background and name
      page.backgrounds = [{ type: 'SOLID', color: { r: 0.961, g: 0.961, b: 0.961 } }]; // #F5F5F5
      page.name = courseName + ' \u2014 Class ' + sessionNum + ': ' + cleanSessionLabel(sessionName);

      // Pre-compute progressive sizes for this session
      var titleGroupMaxLen = {};
      for (var g = 0; g < slides.length; g++) {
        var sl = slides[g];
        if ((sl.type === 'body' || sl.type === 'list') && sl.title) {
          var key = sl.title;
          var len = (sl.body || '').length;
          if (!titleGroupMaxLen[key] || len > titleGroupMaxLen[key]) {
            titleGroupMaxLen[key] = len;
          }
        }
      }
      for (var g2 = 0; g2 < slides.length; g2++) {
        var sl2 = slides[g2];
        if ((sl2.type === 'body' || sl2.type === 'list') && sl2.title && titleGroupMaxLen[sl2.title]) {
          sl2._groupSize = bodySize({ length: titleGroupMaxLen[sl2.title] });
        }
      }

      // Cover slide first
      var coverFrame = buildCoverSlide(courseName, sessionNum, sessionName, 0, 0, logoData, totalSessions);
      page.appendChild(coverFrame);

      // Build frames on this page
      for (var index = 0; index < slides.length; index++) {
        var slide = slides[index];
        slide._courseName = courseName;
        slide._sessionLabel = sessionName;
        if (includePageNumbers) slide._pageNum = index + 2;
        var x = (index + 1) * (W + 80);
        var frame = buildOrCloneFrame(slide, x, 0);
        page.appendChild(frame);
      }

      totalSlides += slides.length + 1; // +1 for cover
      console.log('[PLUGIN] Built session', (i + 1), 'of', sessionGroups.length, '— S' + sessionNum + ' (' + slides.length + ' slides)');
    } catch (sessionErr) {
      console.error('[PLUGIN] ERROR building S' + sessionNum + ':', String(sessionErr));
      errors.push('S' + sessionNum + ': ' + sessionErr.message);
    }

    // Yield to Figma UI thread between sessions to prevent timeout
    await new Promise(function(resolve) { setTimeout(resolve, 300); });
  }

  if (errors.length > 0) {
    figma.ui.postMessage({ type: 'error', message: 'Some sessions had errors: ' + errors.join('; ') });
  }

  // Switch to the first page
  if (sessionGroups.length > 0) {
    figma.currentPage = figma.root.children[figma.root.children.length - totalSessions];
  }

  figma.ui.postMessage({
    type: 'doneAll',
    totalSlides: totalSlides,
    totalSessions: totalSessions
  });
  sendPageCharts();
}

// ============================================================
// FRAME BUILDER
// ============================================================

function buildFrame(slide, x, y) {
  // If title ends with ":" it's a lead-in phrase — merge into body
  if (slide.title && /:\s*$/.test(slide.title) && slide.body) {
    slide.body = slide.title + '\n' + slide.body;
    slide.title = '';
  }

  var frame = figma.createFrame();
  frame.resize(W, H);
  frame.x = x;
  frame.y = y;
  frame.fills = [{ type: 'SOLID', color: COLORS.bg }];
  frame.name = '[' + (slide.type || 'body').toUpperCase() + '] S' + slide.sessionNum + ' \u00B7 ' + slide.number + ' \u2014 ' + (slide.title || 'Untitled');

  switch (slide.type) {
    case 'statement':  buildStatementSlide(frame, slide); break;
    case 'quote':      buildQuoteSlide(frame, slide); break;
    case 'scripture':  buildScriptureSlide(frame, slide); break;
    case 'list':       buildListSlide(frame, slide); break;
    case 'graphic':    buildGraphicSlide(frame, slide); break;
    case 'chart':      buildChartSlide(frame, slide); break;
    case 'illustration': buildIllustrationSlide(frame, slide); break;
    case 'table':      buildTableSlide(frame, slide); break;
    default:           buildBodySlide(frame, slide);
  }

  addFooter(frame, slide._courseName || '', slide._sessionLabel || '');
  if (slide._pageNum) {
    var pgNode = addText(frame, String(slide._pageNum), {
      x: (W - 100) / 2,
      y: FOOTER_Y,
      w: 100,
      h: 32,
      size: FOOTER_SIZE,
      color: COLORS.textFaint,
      opacity: 0.5,
      align: 'CENTER',
      font: 'sans'
    });
    if (pgNode) pgNode.name = '[PAGE_NUM]';
  }
  return frame;
}

// ============================================================
// FOOTER
// ============================================================

function addFooter(frame, courseName, sessionLabel) {
  sessionLabel = cleanSessionLabel(sessionLabel);
  if (courseName) {
    addText(frame, courseName, {
      x: FOOTER_MARGIN,
      y: FOOTER_Y,
      w: 600,
      h: 32,
      size: FOOTER_SIZE,
      color: COLORS.textFaint,
      opacity: 0.5,
      align: 'LEFT',
      font: 'sans'
    });
  }
  if (sessionLabel) {
    addText(frame, sessionLabel, {
      x: W - FOOTER_MARGIN - 600,
      y: FOOTER_Y,
      w: 600,
      h: 32,
      size: FOOTER_SIZE,
      color: COLORS.textFaint,
      opacity: 0.5,
      align: 'RIGHT',
      font: 'sans'
    });
  }
}

// ============================================================
// SLIDE BUILDERS
// ============================================================

function buildStatementSlide(frame, slide) {
  var rawSrc = slide.title || slide.body || '';
  var parts = slide._rawEdit ? { body: rawSrc, attribution: '' } : extractAttribution(rawSrc);
  var text = slide._rawEdit ? parts.body : ensurePeriod(polishText(parts.body));
  text = text.replace(/["""\u201D]+\s*$/, '').trim();
  var attribution = parts.attribution ? cleanAttribution(polishText(parts.attribution)) : '';
  var sz = bodySize(text);
  var refSize = sz - REF_OFFSET;

  var node = addText(frame, text, {
    x: SIDE_MARGIN,
    y: 0,
    w: CONTENT_W,
    h: 200,
    size: sz,
    color: COLORS.textPrimary,
    align: 'CENTER',
    font: 'serif'
  });

  var refNode = null;
  if (attribution) {
    refNode = addText(frame, attribution, {
      x: SIDE_MARGIN,
      y: 0,
      w: CONTENT_W,
      h: 60,
      size: refSize,
      color: COLORS.textMuted,
      opacity: 0.75,
      align: 'CENTER',
      font: 'serif',
      italic: true
    });
  }

  var nodes = node ? [node] : [];
  if (refNode) nodes.push(refNode);
  if (nodes.length > 0) {
    centerBlockVertically(nodes, CONTENT_TOP, CONTENT_BOTTOM);
  }
}

function buildQuoteSlide(frame, slide) {
  // If title is a scripture reference, use it as attribution instead of header
  var scriptureFromTitle = '';
  if (slide.title && isScriptureTitle(slide.title)) {
    scriptureFromTitle = slide.title;
    slide.title = '';
  }

  var parts = slide._rawEdit ? { body: slide.body, attribution: slide.attribution || '' } : extractAttribution(slide.body);
  var body = slide._rawEdit ? parts.body : ensurePeriod(polishText(cleanQuote(parts.body)));
  var attribution = parts.attribution ? (slide._rawEdit ? parts.attribution : cleanAttribution(polishText(parts.attribution))) : '';
  // Use scripture ref from title as attribution if no other attribution found
  if (!attribution && scriptureFromTitle) {
    attribution = slide._rawEdit ? scriptureFromTitle : polishText(scriptureFromTitle);
  }
  var sz = bodySize(body);
  var refSize = sz - REF_OFFSET;

  if (slide.title) {
    addText(frame, slide._rawEdit ? slide.title : polishText(slide.title), {
      x: SIDE_MARGIN,
      y: LABEL_Y,
      w: CONTENT_W,
      h: 40,
      size: TITLE_SIZE,
      color: COLORS.textPrimary,
      align: 'CENTER',
      font: 'sans'
    });
  }

  var bodyNode = addText(frame, body, {
    x: SIDE_MARGIN,
    y: 0,
    w: CONTENT_W,
    h: 200,
    size: sz,
    color: COLORS.textPrimary,
    align: 'CENTER',
    font: 'serif'
  });

  var refNode = null;
  if (attribution) {
    refNode = addText(frame, attribution, {
      x: SIDE_MARGIN,
      y: 0,
      w: CONTENT_W,
      h: 60,
      size: refSize,
      color: COLORS.textMuted,
      opacity: 0.75,
      align: 'CENTER',
      font: 'serif',
      italic: true
    });
  }

  var nodes = bodyNode ? [bodyNode] : [];
  if (refNode) nodes.push(refNode);
  centerBlockVertically(nodes, CONTENT_TOP, CONTENT_BOTTOM);
}

function buildScriptureSlide(frame, slide) {
  var bodyText = slide._rawEdit ? slide.body : polishText(cleanQuote(slide.body));
  var ref = '';

  // If title is already a scripture reference, use it
  if (slide.title && isScriptureTitle(slide.title)) {
    ref = slide._rawEdit ? slide.title : polishText(slide.title);
  } else {
    // Try to extract a scripture reference from the first line of the body
    var lines = bodyText.split('\n');
    for (var li = 0; li < lines.length; li++) {
      var ln = lines[li].trim();
      if (ln && isScriptureTitle(ln)) {
        ref = ln;
        lines.splice(li, 1);
        bodyText = lines.join('\n').trim();
        break;
      }
    }
    // If still no ref, fall back to title as-is
    if (!ref && slide.title) {
      ref = slide._rawEdit ? slide.title : polishText(slide.title);
    }
  }

  // Show title at top if it's not a scripture reference (i.e. it's a real title)
  var showTitle = slide.title && !isScriptureTitle(slide.title);
  if (showTitle) {
    addText(frame, slide._rawEdit ? slide.title : polishText(slide.title), {
      x: SIDE_MARGIN,
      y: LABEL_Y,
      w: CONTENT_W,
      h: 40,
      size: TITLE_SIZE,
      color: COLORS.textPrimary,
      align: 'CENTER',
      font: 'sans'
    });
  }

  var body = bodyText;
  var sz = bodySize(body);
  var refSize = sz - REF_OFFSET;
  var areaTop = showTitle ? CONTENT_TOP : CONTENT_TOP;

  var bodyNode = addText(frame, body, {
    x: SIDE_MARGIN,
    y: 0,
    w: CONTENT_W,
    h: 200,
    size: sz,
    color: COLORS.textPrimary,
    align: 'CENTER',
    font: 'serif'
  });

  var refNode = null;
  if (ref) {
    refNode = addText(frame, ref, {
      x: SIDE_MARGIN,
      y: 0,
      w: CONTENT_W,
      h: 60,
      size: refSize,
      color: COLORS.textMuted,
      opacity: 0.75,
      align: 'CENTER',
      font: 'serif',
      italic: true
    });
  }

  var nodes = bodyNode ? [bodyNode] : [];
  if (refNode) nodes.push(refNode);
  centerBlockVertically(nodes, CONTENT_TOP, CONTENT_BOTTOM);
}

function buildListSlide(frame, slide) {
  // Never show scripture references as list title headers
  if (slide.title && isScriptureTitle(slide.title)) {
    slide.title = '';
  }

  var body = slide._rawEdit ? slide.body : polishText(slide.body);

  if (slide.title) {
    addText(frame, slide._rawEdit ? slide.title : polishText(slide.title), {
      x: SIDE_MARGIN,
      y: LABEL_Y,
      w: CONTENT_W,
      h: 40,
      size: TITLE_SIZE,
      color: COLORS.textPrimary,
      align: 'CENTER',
      font: 'sans'
    });
  }

  // Parse list items
  var rawLines = body.split('\n');
  var items = [];
  for (var li = 0; li < rawLines.length; li++) {
    var line = rawLines[li].trim();
    if (!line) continue;
    // Detect sub-items: SUB_ITEM marker, ○, en dash with indent, or leading whitespace
    var isSub = /^SUB_ITEM\s/.test(line) || /^[○◦]/.test(line) || /^\s{2,}[\u2013\u2014●•\-\*○◦]/.test(rawLines[li]) || /^\s+\u2013/.test(rawLines[li]);
    // Strip all bullet/marker characters (including em/en dashes from prior formatting)
    var clean = line.replace(/^SUB_ITEM\s*/, '').replace(/^[●•○◦\u2014\u2013\-\*]\s*/, '').replace(/^\d+\.\s*/, '').trim();
    if (!clean) continue;
    if (!slide._rawEdit) clean = ensurePeriod(clean);
    items.push({ text: clean, sub: isSub });
  }

  // Check if original had numbered items (1. 2. 3.)
  var hasNumbered = /^\d+\./.test(body.trim());

  // Build formatted list text
  var formatted = items.map(function (item, idx) {
    if (hasNumbered && !item.sub) {
      // Numbered list: keep numbering
      var num = 0;
      for (var ni = 0; ni <= idx; ni++) { if (!items[ni].sub) num++; }
      return num + '.  ' + item.text;
    }
    if (item.sub) {
      return '       \u2013  ' + item.text;  // indented en dash for sub-items
    }
    return '\u2014  ' + item.text;  // em dash for top-level
  }).join('\n');

  // Use the same bodySize scaling as other slide types
  var sz = slide._groupSize || bodySize(formatted);

  var listNode = addText(frame, formatted, {
    x: SIDE_MARGIN + 40,
    y: 0,
    w: CONTENT_W - 80,
    h: 400,
    size: sz,
    color: COLORS.textBody,
    align: 'LEFT',
    font: 'serif',
    lineHeight: 1.9
  });

  if (listNode) {
    clampToZone(listNode, CONTENT_TOP, CONTENT_BOTTOM);
    centerBlockVertically([listNode], CONTENT_TOP, CONTENT_BOTTOM);
  }
}

function buildBodySlide(frame, slide) {
  // If title is a scripture reference, check if the body has a more specific ref
  // (e.g. title "Revelation 4" + body contains "Revelation 4:6")
  // If so, keep title as title and let the body ref be the attribution
  var scriptureFromTitle = '';
  if (slide.title && isScriptureTitle(slide.title)) {
    var bodyHasSpecificRef = false;
    var titleBase = slide.title.replace(/\s+/g, '').toLowerCase();
    var bodyLines = (slide.body || '').split('\n');
    for (var bl = 0; bl < bodyLines.length; bl++) {
      var ln = bodyLines[bl].trim();
      if (isScriptureTitle(ln)) {
        var refBase = ln.replace(/\s+/g, '').toLowerCase();
        // Check if body ref is more specific (has : verse) and shares same book/chapter
        if (refBase.indexOf(':') !== -1 && refBase.indexOf(titleBase) === 0) {
          bodyHasSpecificRef = true;
          break;
        }
      }
    }
    if (bodyHasSpecificRef) {
      // Keep title as title — body's specific ref will become attribution naturally
    } else {
      scriptureFromTitle = slide.title;
      slide.title = '';
    }
  }

  // Detect definition-style content (lines with "=") BEFORE polishText merges lines
  var rawBody = slide.body || '';
  var rawLines = rawBody.split('\n').map(function(l) { return l.trim(); }).filter(function(l) { return l.length > 0; });
  var defLines = rawLines.filter(function(l) { return /\s+=\s+/.test(l); });
  var isDefinitionStyle = defLines.length >= 1 && defLines.length === rawLines.length;

  // Extract scripture references / attributions from body text
  var parts = slide._rawEdit ? { body: slide.body, attribution: '' } : extractAttribution(slide.body);
  var body;

  if (slide._rawEdit) {
    // Raw edit mode — user controls punctuation and formatting directly
    body = parts.body;
  } else if (isDefinitionStyle) {
    // For definition content, skip polishText line-merging — format each line individually
    body = rawLines.map(function(line) {
      var eqIdx = line.indexOf('=');
      if (eqIdx > 0) {
        var left = line.substring(0, eqIdx).trim();
        var right = line.substring(eqIdx + 1).trim();
        return '*' + left + '* = ' + right;
      }
      return line;
    }).join('\n');
  } else {
    body = polishText(parts.body);
    // Strip any "Static Graphic" lines that leaked into body content
    body = body.replace(/\n?Static\s*(Graphic|Slide)\s*(\([\d:]*\))?\s*/gi, '').trim();
    body = ensurePeriod(body);
  }
  // Also clean trailing quote marks left after ref extraction
  body = body.replace(/["""\u201D]+\s*$/, '').trim();
  var attribution = parts.attribution ? cleanAttribution(polishText(parts.attribution)) : '';
  // Use scripture ref from title as attribution if no other attribution found
  if (!attribution && scriptureFromTitle) {
    attribution = slide._rawEdit ? scriptureFromTitle : polishText(scriptureFromTitle);
  }
  var sz = slide._groupSize || bodySize(body);
  var refSize = sz - REF_OFFSET;

  if (slide.title) {
    addText(frame, slide._rawEdit ? slide.title : polishText(slide.title), {
      x: SIDE_MARGIN,
      y: LABEL_Y,
      w: CONTENT_W,
      h: 40,
      size: TITLE_SIZE,
      color: COLORS.textPrimary,
      align: 'CENTER',
      font: 'sans'
    });
  }

  var bodyNode = addText(frame, body, {
    x: SIDE_MARGIN,
    y: 0,
    w: CONTENT_W,
    h: 200,
    size: sz,
    color: COLORS.textPrimary,
    align: 'CENTER',
    font: 'serif',
    lineHeight: 1.8
  });

  var refNode = null;
  if (attribution) {
    refNode = addText(frame, attribution, {
      x: SIDE_MARGIN,
      y: 0,
      w: CONTENT_W,
      h: 60,
      size: refSize,
      color: COLORS.textMuted,
      opacity: 0.75,
      align: 'CENTER',
      font: 'serif',
      italic: true
    });
  }

  var nodes = bodyNode ? [bodyNode] : [];
  if (refNode) nodes.push(refNode);
  if (nodes.length > 0) {
    for (var n = 0; n < nodes.length; n++) {
      clampToZone(nodes[n], CONTENT_TOP, CONTENT_BOTTOM);
    }
    centerBlockVertically(nodes, CONTENT_TOP, CONTENT_BOTTOM);
  }
}

// Title/holding slide — "Static Graphic" in the docs
// Empty placeholder — just the background + footer (user adds art manually)
function buildGraphicSlide(frame, slide) {
  // Intentionally empty — only footer is shown (added by buildFrame)
}

// Cover slide — logo + course title (Signifier Medium 96px) centered with class name below,
// "Class X" label near bottom center
function buildCoverSlide(courseName, sessionNum, sessionName, x, y, logoData, totalSessions) {
  var frame = figma.createFrame();
  frame.resize(W, H);
  frame.x = x;
  frame.y = y;
  frame.fills = [{ type: 'SOLID', color: COLORS.bg }];
  frame.name = '[COVER] S' + sessionNum + ' \u2014 ' + cleanSessionLabel(sessionName);

  // Logo above title — rounded square background + SVG icon
  var logoGroup = null;
  if (logoData && logoData.svg) {
    try {
      var LOGO_SIZE = 100; // square side length
      var ICON_SIZE = 60;  // SVG icon size inside square
      var LOGO_RADIUS = 20; // corner radius

      // Create rounded square background
      var rect = figma.createRectangle();
      rect.resize(LOGO_SIZE, LOGO_SIZE);
      rect.cornerRadius = LOGO_RADIUS;
      rect.fills = [{ type: 'SOLID', color: logoData.bgColor ? hexToRgb(logoData.bgColor) : COLORS.textPrimary }];
      rect.x = (W - LOGO_SIZE) / 2;
      rect.y = 0;
      frame.appendChild(rect);

      // Create SVG icon
      var svgFrame = figma.createNodeFromSvg(logoData.svg);
      var scale = ICON_SIZE / 54;
      var iconW = Math.round(55 * scale);
      var iconH = ICON_SIZE;
      svgFrame.resize(iconW, iconH);
      svgFrame.x = (W - iconW) / 2;
      svgFrame.y = (LOGO_SIZE - iconH) / 2;
      frame.appendChild(svgFrame);

      // Group them for vertical centering
      var grp = figma.group([rect, svgFrame], frame);
      grp.name = 'Logo';
      logoGroup = grp;
    } catch (e) {
      logoGroup = null;
    }
  }

  // Course title — Signifier Medium 96px, centered
  var courseNode = addText(frame, courseName, {
    x: SIDE_MARGIN,
    y: 0,
    w: CONTENT_W,
    h: 120,
    size: 96,
    color: COLORS.textPrimary,
    align: 'CENTER',
    font: 'serif',
    weight: 'MEDIUM'
  });

  // Class name — Untitled Sans 36px underneath
  var classLabel = cleanSessionLabel(sessionName);
  var classNode = addText(frame, classLabel, {
    x: SIDE_MARGIN,
    y: 0,
    w: CONTENT_W,
    h: 60,
    size: 36,
    color: COLORS.textPrimary,
    opacity: 0.6,
    align: 'CENTER',
    font: 'sans'
  });

  // Center logo + course title + class name vertically with tight gaps
  var coverNodes = [];
  if (logoGroup) coverNodes.push(logoGroup);
  if (courseNode) coverNodes.push(courseNode);
  if (classNode) coverNodes.push(classNode);
  if (coverNodes.length > 0) {
    // Use a tighter gap (24px) between elements on cover
    var COVER_GAP = 24;
    var totalH = 0;
    for (var i = 0; i < coverNodes.length; i++) {
      totalH += coverNodes[i].height;
      if (i > 0) totalH += COVER_GAP;
    }
    var zone = CONTENT_BOTTOM - CONTENT_TOP;
    var startY = CONTENT_TOP + Math.max(0, (zone - totalH) / 2);
    var cy = startY;
    for (var j = 0; j < coverNodes.length; j++) {
      coverNodes[j].y = cy;
      // Center logo horizontally (text nodes already positioned via addText x)
      if (coverNodes[j] === logoGroup) {
        coverNodes[j].x = (W - coverNodes[j].width) / 2;
      }
      cy += coverNodes[j].height + COVER_GAP;
    }
  }

  // "Class X of Y" label — Untitled Sans 36px, positioned above footer area
  var classLabel2 = totalSessions ? 'Class ' + sessionNum + ' of ' + totalSessions : 'Class ' + sessionNum;
  addText(frame, classLabel2, {
    x: SIDE_MARGIN,
    y: FOOTER_Y - 50,
    w: CONTENT_W,
    h: 40,
    size: 36,
    color: COLORS.textPrimary,
    opacity: 0.5,
    align: 'CENTER',
    font: 'sans'
  });

  return frame;
}

// Chart/illustration placeholder — red marker so you know to add manually
function buildChartSlide(frame, slide) {
  var label = polishText(slide.title || slide.body || 'Chart / Illustration');

  // Red dot marker — top right corner
  var dot = figma.createEllipse();
  dot.resize(24, 24);
  dot.x = W - FOOTER_MARGIN - 24;
  dot.y = 32;
  dot.fills = [{ type: 'SOLID', color: COLORS.redMarker }];
  frame.appendChild(dot);

  // "Illustration Needed" label next to dot
  addText(frame, 'Illustration Needed', {
    x: W - FOOTER_MARGIN - 220,
    y: 34,
    w: 190,
    h: 24,
    size: 14,
    color: COLORS.redMarker,
    align: 'RIGHT',
    font: 'sans'
  });

  // Show whatever context was in the slide
  var bodyNode = addText(frame, label, {
    x: SIDE_MARGIN,
    y: 0,
    w: CONTENT_W,
    h: 200,
    size: 32,
    color: COLORS.textPrimary,
    opacity: 0.4,
    align: 'CENTER',
    font: 'serif',
    italic: true
  });

  if (bodyNode) {
    centerBlockVertically([bodyNode], CONTENT_TOP, CONTENT_BOTTOM);
  }
}

// Illustration/Image slide — renders an image centered with optional title above
// If no image data provided, shows a gold placeholder
function buildIllustrationSlide(frame, slide) {
  var title = polishText(slide.title || '');
  var IMG_AREA_TOP = CONTENT_TOP;
  var IMG_AREA_BOTTOM = CONTENT_BOTTOM;

  // Title above image if present
  if (title) {
    addText(frame, title, {
      x: SIDE_MARGIN,
      y: CONTENT_TOP,
      w: CONTENT_W,
      h: 50,
      size: 28,
      color: COLORS.textPrimary,
      align: 'CENTER',
      font: 'sans'
    });
    IMG_AREA_TOP = CONTENT_TOP + 60;
  }

  var imgW = CONTENT_W - 120; // padding on sides
  var imgH = IMG_AREA_BOTTOM - IMG_AREA_TOP - 40;
  var imgX = SIDE_MARGIN + 60;
  var imgY = IMG_AREA_TOP + 20;

  // If image data was provided (base64), render it
  if (slide._imageData) {
    try {
      var imgBytes = figma.base64Decode(slide._imageData);
      var imgHash = figma.createImage(imgBytes).hash;
      var rect = figma.createRectangle();
      rect.resize(imgW, imgH);
      rect.x = imgX;
      rect.y = imgY;
      rect.fills = [{ type: 'IMAGE', imageHash: imgHash, scaleMode: 'FIT' }];
      rect.cornerRadius = 8;
      frame.appendChild(rect);
    } catch (e) {
      // Fallback to placeholder on error
      renderIllustrationPlaceholder(frame, imgX, imgY, imgW, imgH);
    }
  } else {
    renderIllustrationPlaceholder(frame, imgX, imgY, imgW, imgH);
  }

  // Store image ref in plugin data if present
  if (slide.imageRef) {
    frame.setPluginData('imageRef', slide.imageRef);
  }
}

function renderIllustrationPlaceholder(frame, x, y, w, h) {
  // Dashed border rectangle
  var rect = figma.createRectangle();
  rect.resize(w, h);
  rect.x = x;
  rect.y = y;
  rect.fills = [{ type: 'SOLID', color: COLORS.bg, opacity: 0.3 }];
  rect.strokes = [{ type: 'SOLID', color: { r: 0.831, g: 0.659, b: 0.325 } }]; // #D4A853
  rect.strokeWeight = 2;
  rect.dashPattern = [12, 8];
  rect.cornerRadius = 8;
  frame.appendChild(rect);

  // Gold icon placeholder — camera/image icon
  var iconSize = 48;
  addText(frame, '\u25A3', { // ▣ square with fill
    x: x + (w - iconSize) / 2,
    y: y + (h / 2) - 40,
    w: iconSize,
    h: iconSize,
    size: 40,
    color: { r: 0.831, g: 0.659, b: 0.325 }, // #D4A853
    align: 'CENTER',
    font: 'sans'
  });

  // "Image Placeholder" label
  addText(frame, 'Image Placeholder', {
    x: x,
    y: y + (h / 2) + 10,
    w: w,
    h: 30,
    size: 18,
    color: { r: 0.831, g: 0.659, b: 0.325 }, // #D4A853
    opacity: 0.8,
    align: 'CENTER',
    font: 'sans'
  });
}

function buildTableSlide(frame, slide) {
  var rawBody = slide.body || '';
  var lines = rawBody.split('\n').map(function (l) { return l.trim(); });

  // Filter out stray page numbers and empties
  var cleanLines = lines.filter(function (l) { return l.length > 0 && !/^\d{1,3}$/.test(l); });

  // Identify subtitle, header, and data lines
  var subtitle = '';
  var headerLine = '';
  var dataStartIdx = 0;

  if (cleanLines.length > 0 && !/\d+:\d+/.test(cleanLines[0])) {
    subtitle = cleanLines[0];
    if (cleanLines.length > 1 && !/\d+:\d+/.test(cleanLines[1])) {
      headerLine = cleanLines[1];
      dataStartIdx = 2;
    } else {
      dataStartIdx = 1;
    }
  }

  // Parse header into 4 columns using Bible book names as delimiters
  var bookNames = /\b(Genesis|Exodus|Leviticus|Numbers|Deuteronomy|Joshua|Judges|Ruth|Samuel|Kings|Chronicles|Ezra|Nehemiah|Esther|Job|Psalms?|Proverbs|Ecclesiastes|Song|Isaiah|Jeremiah|Lamentations|Ezekiel|Daniel|Hosea|Joel|Amos|Obadiah|Jonah|Micah|Nahum|Habakkuk|Zephaniah|Haggai|Zechariah|Malachi|Matthew|Mark|Luke|John|Acts|Romans|Corinthians|Galatians|Ephesians|Philippians|Colossians|Thessalonians|Timothy|Titus|Philemon|Hebrews|James|Peter|Jude|Revelation)\b/gi;
  var headers = ['', '', '', ''];
  if (headerLine) {
    var bookMatches = [];
    var bm;
    while ((bm = bookNames.exec(headerLine)) !== null) {
      bookMatches.push({ text: bm[1], index: bm.index, end: bm.index + bm[0].length });
    }
    if (bookMatches.length >= 2) {
      headers[0] = headerLine.substring(0, bookMatches[0].index).trim();
      headers[1] = bookMatches[0].text;
      headers[2] = headerLine.substring(bookMatches[0].end, bookMatches[1].index).trim();
      headers[3] = bookMatches[1].text;
    } else {
      var words = headerLine.split(/\s+/);
      var mid = Math.ceil(words.length / 2);
      headers[0] = words.slice(0, mid).join(' ');
      headers[2] = words.slice(mid).join(' ');
    }
  }

  // Verse ref pattern — handles "3:8-11", "22:2, 14", "19:11-21; 20:7-10"
  var refPattern = /(\d+:\d+(?:[–\-]\d+(?::\d+)?)?(?:[;,]\s*\d+(?::\d+)?(?:[–\-]\d+(?::\d+)?)?)*)/g;

  // Parse a single ref-containing line into 4 cells + trailing info
  function parseRefLine(line) {
    var refs = [];
    var m;
    refPattern.lastIndex = 0;
    while ((m = refPattern.exec(line)) !== null) {
      refs.push({ text: m[1], index: m.index, end: m.index + m[0].length });
    }
    var result = {
      leftText: '', leftRef: '', rightText: '', rightRef: '',
      _leftRefTrailing: false, _rightRefTrailing: false
    };
    if (refs.length >= 2) {
      result.leftText = line.substring(0, refs[0].index).trim();
      result.leftRef = refs[0].text;
      // Check if char after leftRef is ";" (trailing continuation marker)
      var afterLeft = line.substring(refs[0].end);
      result._leftRefTrailing = /^[;,]/.test(afterLeft.trim());
      // rightText: strip leading semicolons/commas/spaces between the refs
      result.rightText = line.substring(refs[0].end, refs[1].index).replace(/^[;,\s]+/, '').trim();
      result.rightRef = refs[1].text;
      // Append any additional refs to rightRef
      for (var rx = 2; rx < refs.length; rx++) {
        result.rightRef += '; ' + refs[rx].text;
      }
      // Check trailing after last ref
      var afterLast = line.substring(refs[refs.length - 1].end);
      result._rightRefTrailing = /^[;,]/.test(afterLast.trim());
    } else if (refs.length === 1) {
      result.leftText = line.substring(0, refs[0].index).trim();
      result.leftRef = refs[0].text;
      result.rightText = line.substring(refs[0].end).replace(/^[;,\s]+/, '').trim();
      var afterOnly = line.substring(refs[0].end);
      result._leftRefTrailing = /^[;,]/.test(afterOnly.trim());
    } else {
      result.leftText = line;
    }
    return result;
  }

  // Determine if text looks like a complete sentence ending
  function endsComplete(s) {
    return /[.!?\u201D\u2019"'\u201C]$/.test(s.trim());
  }

  // Build parsed rows with proper continuation handling
  var parsedRows = [];
  for (var i = dataStartIdx; i < cleanLines.length; i++) {
    var ln = cleanLines[i];
    var hasRef = /\d+:\d+/.test(ln);
    var isRefOnly = hasRef && /^[\d:;\u2013\-,\s]+$/.test(ln);

    // A new row has significant descriptive text (>5 chars) before the first ref
    var isNewRow = false;
    if (hasRef && !isRefOnly) {
      var firstRefIdx = ln.search(/\d+:\d+/);
      if (firstRefIdx > 5) isNewRow = true;
    }

    if (isNewRow) {
      parsedRows.push(parseRefLine(ln));
    } else if (parsedRows.length > 0) {
      var prev = parsedRows[parsedRows.length - 1];
      if (isRefOnly) {
        // Ref-only continuation — append to whichever ref had a trailing ";"
        var refVal = ln.trim();
        if (prev._rightRefTrailing) {
          prev.rightRef = prev.rightRef + '; ' + refVal;
          prev._rightRefTrailing = false;
        } else if (prev._leftRefTrailing) {
          prev.leftRef = prev.leftRef + '; ' + refVal;
          prev._leftRefTrailing = false;
        } else {
          prev.rightRef = (prev.rightRef ? prev.rightRef + '; ' : '') + refVal;
        }
      } else {
        // Text continuation — decide left vs right
        var leftDone = endsComplete(prev.leftText);
        var rightDone = endsComplete(prev.rightText);
        if (!leftDone && !rightDone) {
          // Both sides incomplete — try splitting at a sentence boundary
          var splitMatch = ln.match(/^(.+?[.!?\u201D\u2019"'])\s+(.+)$/);
          if (splitMatch) {
            prev.leftText += ' ' + splitMatch[1];
            prev.rightText += ' ' + splitMatch[2];
          } else {
            prev.leftText += ' ' + ln;
          }
        } else if (!rightDone && leftDone) {
          prev.rightText += ' ' + ln;
        } else {
          prev.leftText += ' ' + ln;
        }
      }
    }
  }

  // ── Layout — centered on slide ──
  var TABLE_W = 1560;
  var TABLE_X = (W - TABLE_W) / 2;
  var HEADER_H = 52;
  var LINE_THICK = 1.5;
  var CELL_PAD_X = 20;
  var CELL_PAD_Y = 12;
  var TEXT_SIZE = 20;
  var REF_SIZE = 17;
  var HEADER_SIZE = 18;

  // Column widths: leftText | leftRef | rightText | rightRef
  var refColW = Math.round(TABLE_W * 0.09);
  var textColW = Math.round((TABLE_W - refColW * 2) / 2);
  var colWidths = [textColW, refColW, textColW, refColW];

  // Measure row heights
  var charsPerLine = Math.floor((textColW - CELL_PAD_X * 2) / (TEXT_SIZE * 0.5));
  var rowHeights = [];
  var totalDataH = 0;
  for (var rh = 0; rh < parsedRows.length; rh++) {
    var leftLines = Math.ceil((parsedRows[rh].leftText || '').length / charsPerLine) || 1;
    var rightLines = Math.ceil((parsedRows[rh].rightText || '').length / charsPerLine) || 1;
    var maxLines = Math.max(leftLines, rightLines);
    var rowH = maxLines * (TEXT_SIZE * 1.55) + CELL_PAD_Y * 2;
    rowH = Math.max(rowH, 56);
    rowHeights.push(rowH);
    totalDataH += rowH;
  }

  // Scale if too tall
  var titleBlockH = 120; // title + subtitle space
  var maxTableH = FOOTER_Y - titleBlockH - 40;
  if (totalDataH + HEADER_H > maxTableH) {
    var scale = (maxTableH - HEADER_H) / totalDataH;
    totalDataH = 0;
    for (var rs = 0; rs < rowHeights.length; rs++) {
      rowHeights[rs] = Math.max(46, Math.round(rowHeights[rs] * scale));
      totalDataH += rowHeights[rs];
    }
  }

  var totalHeight = HEADER_H + totalDataH;

  // Title at standard top position, subtitle below it
  var titleY = CONTENT_TOP;
  var subtitleY = CONTENT_TOP + 50;
  var tableAreaTop = subtitle ? subtitleY + 40 : titleY + 50;

  // Center the table vertically in the remaining space between title block and footer
  var tableAreaH = FOOTER_Y - tableAreaTop - 20;
  var TABLE_TOP = tableAreaTop + Math.round((tableAreaH - totalHeight) / 2);
  if (TABLE_TOP < tableAreaTop) TABLE_TOP = tableAreaTop;

  // ── Title ──
  if (slide.title) {
    addText(frame, slide._rawEdit ? slide.title : polishText(slide.title), {
      x: SIDE_MARGIN, y: titleY, w: CONTENT_W, h: 40,
      size: TITLE_SIZE, color: COLORS.textPrimary, align: 'CENTER', font: 'sans'
    });
  }

  // ── Subtitle ──
  if (subtitle) {
    addText(frame, subtitle, {
      x: SIDE_MARGIN, y: subtitleY, w: CONTENT_W, h: 32,
      size: 24, color: COLORS.textPrimary, opacity: 0.55,
      align: 'CENTER', font: 'serif', italic: true
    });
  }

  // ── Line helpers ──
  function drawHLine(yPos) {
    var line = figma.createRectangle();
    line.resize(TABLE_W, LINE_THICK);
    line.x = TABLE_X;
    line.y = yPos;
    line.fills = [{ type: 'SOLID', color: COLORS.textPrimary }];
    line.opacity = 0.15;
    frame.appendChild(line);
  }

  function drawVLine(xPos, yStart, height) {
    var line = figma.createRectangle();
    line.resize(LINE_THICK, height);
    line.x = xPos;
    line.y = yStart;
    line.fills = [{ type: 'SOLID', color: COLORS.textPrimary }];
    line.opacity = 0.15;
    frame.appendChild(line);
  }

  // ── Inner lines ──
  drawHLine(TABLE_TOP + HEADER_H);
  var rowY = TABLE_TOP + HEADER_H;
  for (var rd = 0; rd < parsedRows.length - 1; rd++) {
    rowY += rowHeights[rd];
    drawHLine(rowY);
  }

  // Center vertical divider
  var centerX = TABLE_X + colWidths[0] + colWidths[1];
  drawVLine(centerX, TABLE_TOP, totalHeight);

  // ── 4-column headers ──
  var hY = TABLE_TOP + (HEADER_H - HEADER_SIZE * 1.3) / 2;
  addText(frame, headers[0], {
    x: TABLE_X + CELL_PAD_X, y: hY,
    w: colWidths[0] - CELL_PAD_X * 2, h: HEADER_H - 12,
    size: HEADER_SIZE, color: COLORS.textPrimary, opacity: 0.6,
    align: 'LEFT', font: 'sans'
  });
  addText(frame, headers[1], {
    x: TABLE_X + colWidths[0] + 6, y: hY,
    w: colWidths[1] - 12, h: HEADER_H - 12,
    size: HEADER_SIZE, color: COLORS.textPrimary, opacity: 0.6,
    align: 'LEFT', font: 'sans'
  });
  addText(frame, headers[2], {
    x: centerX + CELL_PAD_X, y: hY,
    w: colWidths[2] - CELL_PAD_X * 2, h: HEADER_H - 12,
    size: HEADER_SIZE, color: COLORS.textPrimary, opacity: 0.6,
    align: 'LEFT', font: 'sans'
  });
  addText(frame, headers[3], {
    x: centerX + colWidths[2] + 6, y: hY,
    w: colWidths[3] - 12, h: HEADER_H - 12,
    size: HEADER_SIZE, color: COLORS.textPrimary, opacity: 0.6,
    align: 'LEFT', font: 'sans'
  });

  // ── Data rows ──
  rowY = TABLE_TOP + HEADER_H;
  for (var ri = 0; ri < parsedRows.length; ri++) {
    var pr = parsedRows[ri];
    var cellY = rowY + CELL_PAD_Y;
    var cellH = rowHeights[ri] - CELL_PAD_Y * 2;

    addText(frame, pr.leftText, {
      x: TABLE_X + CELL_PAD_X, y: cellY,
      w: colWidths[0] - CELL_PAD_X * 2, h: cellH,
      size: TEXT_SIZE, color: COLORS.textPrimary,
      align: 'LEFT', font: 'serif'
    });
    addText(frame, pr.leftRef, {
      x: TABLE_X + colWidths[0] + 6, y: cellY,
      w: colWidths[1] - 12, h: cellH,
      size: REF_SIZE, color: COLORS.textPrimary, opacity: 0.5,
      align: 'LEFT', font: 'sans'
    });
    addText(frame, pr.rightText, {
      x: centerX + CELL_PAD_X, y: cellY,
      w: colWidths[2] - CELL_PAD_X * 2, h: cellH,
      size: TEXT_SIZE, color: COLORS.textPrimary,
      align: 'LEFT', font: 'serif'
    });
    addText(frame, pr.rightRef, {
      x: centerX + colWidths[2] + 6, y: cellY,
      w: colWidths[3] - 12, h: cellH,
      size: REF_SIZE, color: COLORS.textPrimary, opacity: 0.5,
      align: 'LEFT', font: 'sans'
    });

    rowY += rowHeights[ri];
  }
}

// ============================================================
// CLAMP — scale down font if content overflows the zone
// ============================================================

function clampToZone(node, top, bottom) {
  var zone = bottom - top;
  while (node.height > zone && node.fontSize > 28) {
    node.fontSize = node.fontSize - 4;
    node.lineHeight = { value: node.fontSize * 1.55, unit: 'PIXELS' };
  }
}

// ============================================================
// VERTICAL CENTERING — positions a block of nodes centered
// between top and bottom bounds, with REF_GAP between them
// ============================================================

function centerBlockVertically(nodes, top, bottom) {
  if (nodes.length === 0) return;

  // Calculate total block height: sum of node heights + gaps between
  var totalH = 0;
  for (var i = 0; i < nodes.length; i++) {
    totalH += nodes[i].height;
    if (i > 0) totalH += REF_GAP;
  }

  // Center the block in the available zone
  var zone = bottom - top;
  var startY = top + Math.max(0, (zone - totalH) / 2);

  var y = startY;
  for (var j = 0; j < nodes.length; j++) {
    nodes[j].y = y;
    y += nodes[j].height + REF_GAP;
  }
}

// ============================================================
// TEXT NODE HELPER
// ============================================================

function addText(frame, content, opts) {
  if (!content || content.trim() === '') return null;

  // Parse *emphasis* markers — collect ranges before stripping markers
  var emphRanges = [];
  var parsed = parseEmphasis(content);
  var cleanContent = parsed.text;
  emphRanges = parsed.ranges;

  var node = figma.createText();
  frame.appendChild(node);

  // Determine font: check role overrides first, then fall back to defaults
  var roleName = opts.role || null;
  var fontName = null;

  // Auto-detect role from opts if not explicitly set
  if (!roleName) {
    if (opts.font === 'sans') roleName = 'labels';
    else if (opts.italic) roleName = 'attribution';
    else if (opts.weight === 'BOLD' || opts.weight === 'MEDIUM') roleName = 'title';
    else roleName = 'body';
  }

  // Use role override if set
  if (FONT_ROLES[roleName]) {
    fontName = FONT_ROLES[roleName];
  } else {
    // Default key lookup
    var fontKey;
    if (opts.font === 'sans') {
      fontKey = opts.weight === 'BOLD' ? 'sansBold' : (opts.weight === 'MEDIUM' ? 'sansMedium' : 'sansRegular');
    } else {
      if (opts.italic) {
        fontKey = 'serifItalic';
      } else if (opts.weight === 'BOLD' || opts.weight === 'MEDIUM') {
        fontKey = 'serifMedium';
      } else {
        fontKey = 'serifRegular';
      }
    }
    fontName = RESOLVED_FONTS[fontKey];
  }

  node.fontName = fontName;
  node.characters = cleanContent;
  node.fontSize = opts.size || 24;
  node.fills = [{ type: 'SOLID', color: opts.color || COLORS.textPrimary }];
  if (typeof opts.opacity === 'number') {
    node.opacity = opts.opacity;
  }
  node.textAlignHorizontal = opts.align || 'LEFT';
  node.textAlignVertical = 'TOP';
  node.resize(opts.w, opts.h);
  node.x = opts.x;
  node.y = opts.y;

  var lh = opts.lineHeight || 1.55;
  node.lineHeight = { value: opts.size * lh, unit: 'PIXELS' };

  // Apply emphasis ranges — use role override or default Medium Italic
  var emphFont = FONT_ROLES.emphasis || RESOLVED_FONTS.serifMediumItalic;
  if (emphRanges.length > 0 && emphFont) {
    for (var e = 0; e < emphRanges.length; e++) {
      var range = emphRanges[e];
      node.setRangeFontName(range.start, range.end, emphFont);
    }
  }

  // Auto-resize height so we can measure actual content height
  node.textAutoResize = 'HEIGHT';

  return node;
}

// Parse *emphasis* markers from text. Returns { text, ranges[] }
// Supports *single asterisks* for emphasis.
// Ranges are [start, end) indices into the cleaned (marker-free) text.
function parseEmphasis(text) {
  var ranges = [];
  var out = '';
  var i = 0;
  while (i < text.length) {
    if (text[i] === '*' && i + 1 < text.length && text[i + 1] !== '*' && text[i + 1] !== ' ') {
      // Opening * — find closing *
      var closeIdx = text.indexOf('*', i + 1);
      if (closeIdx > i + 1 && text[closeIdx - 1] !== ' ') {
        // Valid emphasis span
        var start = out.length;
        var inner = text.substring(i + 1, closeIdx);
        out += inner;
        ranges.push({ start: start, end: out.length });
        i = closeIdx + 1;
        continue;
      }
    }
    out += text[i];
    i++;
  }
  return { text: out, ranges: ranges };
}

// ============================================================
// UTILITIES
// ============================================================

var BOOK_ABBREVS = {
  'Gen\\.?': 'Genesis', 'Exod\\.?': 'Exodus', 'Lev\\.?': 'Leviticus',
  'Num\\.?': 'Numbers', 'Deut\\.?': 'Deuteronomy', 'Josh\\.?': 'Joshua',
  'Judg\\.?': 'Judges', 'Sam\\.?': 'Samuel', 'Kgs\\.?': 'Kings',
  'Chr\\.?': 'Chronicles', 'Neh\\.?': 'Nehemiah', 'Esth\\.?': 'Esther',
  'Ps\\.?': 'Psalm', 'Pss\\.?': 'Psalms', 'Prov\\.?': 'Proverbs',
  'Eccl\\.?': 'Ecclesiastes', 'Isa\\.?': 'Isaiah', 'Jer\\.?': 'Jeremiah',
  'Lam\\.?': 'Lamentations', 'Ezek\\.?': 'Ezekiel', 'Dan\\.?': 'Daniel',
  'Hos\\.?': 'Hosea', 'Obad\\.?': 'Obadiah', 'Mic\\.?': 'Micah',
  'Nah\\.?': 'Nahum', 'Hab\\.?': 'Habakkuk', 'Zeph\\.?': 'Zephaniah',
  'Hag\\.?': 'Haggai', 'Zech\\.?': 'Zechariah', 'Mal\\.?': 'Malachi',
  'Matt\\.?': 'Matthew', 'Rom\\.?': 'Romans', 'Cor\\.?': 'Corinthians',
  'Gal\\.?': 'Galatians', 'Eph\\.?': 'Ephesians', 'Phil\\.?': 'Philippians',
  'Col\\.?': 'Colossians', 'Thess\\.?': 'Thessalonians', 'Tim\\.?': 'Timothy',
  'Phlm\\.?': 'Philemon', 'Heb\\.?': 'Hebrews', 'Jas\\.?': 'James',
  'Pet\\.?': 'Peter', 'Rev\\.?': 'Revelation'
};

function expandBookNames(text) {
  var s = text;
  var abbrevs = Object.keys(BOOK_ABBREVS);
  for (var i = 0; i < abbrevs.length; i++) {
    var pattern = new RegExp('\\b(\\d\\s*)?' + abbrevs[i] + '(?=\\s+\\d)', 'g');
    var full = BOOK_ABBREVS[abbrevs[i]];
    s = s.replace(pattern, function (match, prefix) {
      return (prefix || '') + full;
    });
  }
  return s;
}

// Strip version numbers from session labels (e.g. "Salvation Pt 1 - V2" → "Salvation, Part 1")
function cleanSessionLabel(label) {
  var s = label;
  // Remove version markers: " - V2", " -V4", " V2", "- V2" etc.
  s = s.replace(/\s*[-–—]\s*V\d+\s*$/i, '');
  s = s.replace(/\s+V\d+\s*$/i, '');
  // Clean up "Pt" → "Part"
  s = s.replace(/\bPt\s*(\d)/g, 'Part $1');
  return s.trim();
}

function cleanQuote(text) {
  // Strip leading open-quotes, trailing close-quotes, and trailing dashes
  return text
    .replace(/^[\u201C\u2018"']+\s*/g, '')
    .replace(/[\u201D\u2019"']+\s*$/g, '')
    .replace(/\s*[—–\-]+\s*$/g, '')
    .trim();
}

function extractAttribution(text) {
  var s = text.trim();

  // Normalize non-breaking spaces to regular spaces for regex matching
  s = s.replace(/\u00A0/g, ' ');

  // Collapse newlines inside parentheses (PDF wraps like "(Lev.\n17:11)")
  s = s.replace(/\(([^)]*)\)/g, function (match, inner) {
    return '(' + inner.replace(/\n/g, ' ') + ')';
  });

  // 1. Parenthetical scripture reference at end: ...(John 1:1-5) or ...(1 Thessalonians 1:4-5a)
  //    Uses comprehensive Bible book name detection
  var bookNames = 'Genesis|Exodus|Leviticus|Numbers|Deuteronomy|Joshua|Judges|Ruth|Samuel|Kings|Chronicles|Ezra|Nehemiah|Esther|Job|Psalms?|Proverbs|Ecclesiastes|Song|Isaiah|Jeremiah|Lamentations|Ezekiel|Daniel|Hosea|Joel|Amos|Obadiah|Jonah|Micah|Nahum|Habakkuk|Zephaniah|Haggai|Zechariah|Malachi|Matthew|Mark|Luke|John|Acts|Romans|Corinthians|Galatians|Ephesians|Philippians|Colossians|Thessalonians|Timothy|Titus|Philemon|Hebrews|James|Peter|Jude|Revelation|Gen|Exod|Lev|Num|Deut|Josh|Judg|Sam|Kgs|Chr|Neh|Esth|Ps|Pss|Prov|Eccl|Isa|Jer|Lam|Ezek|Dan|Hos|Obad|Mic|Nah|Hab|Zeph|Hag|Zech|Mal|Matt|Rom|Cor|Gal|Eph|Phil|Col|Thess|Tim|Phlm|Heb|Jas|Pet|Rev';
  var parenRefRegex = new RegExp('\\(([123]?\\s*(?:' + bookNames + ')\\.?\\s+\\d[\\d:,\\-–;\\sa-z]*)\\)\\.?\\s*$', 'i');
  var parenRefMatch = s.match(parenRefRegex);
  if (parenRefMatch) {
    var beforeParen = s.substring(0, s.lastIndexOf(parenRefMatch[0])).trim();
    // Only extract as attribution if the text before ends with a closing quote mark,
    // meaning this is a quote attribution. If it's regular body text (no closing quote),
    // the scripture ref is an inline supporting reference — keep it in place.
    if (/["""\u201D\u2019']\s*$/.test(beforeParen)) {
      beforeParen = beforeParen.replace(/["""\u201D]+\s*$/, '');
      return {
        body: beforeParen,
        attribution: parenRefMatch[1].trim()
      };
    }
    // Not a quote — leave the parenthetical reference inline
  }

  // 2. Inline attribution: ...quote text" - Author Name
  var inlineMatch = s.match(/(["""\u201D]\.?)\s*[-—–]\s*([A-Z][\w\s.,]+)$/);
  if (inlineMatch) {
    var bodyEnd = s.lastIndexOf(inlineMatch[0]);
    var quoteBody = s.substring(0, bodyEnd).trim();
    // Strip trailing quote marks from body
    quoteBody = quoteBody.replace(/["""\u201D]+\s*$/, '');
    return {
      body: quoteBody,
      attribution: inlineMatch[2].trim()
    };
  }

  // 2b. Scripture combo: context line(s) + scripture ref line + verse text
  // e.g. "John gets a vision of the Throne Room.\nRevelation 4:6\n"and before the throne..."
  // The scripture ref line becomes attribution, everything else stays as body
  var comboLines = s.split('\n');
  for (var ci = 0; ci < comboLines.length; ci++) {
    var cln = comboLines[ci].trim();
    if (isScriptureTitle(cln) && cln.indexOf(':') !== -1) {
      // Found a scripture reference line in the middle/start of body
      var beforeRef = comboLines.slice(0, ci).join('\n').trim();
      var afterRef = comboLines.slice(ci + 1).join('\n').trim();
      // Combine context + verse as body, use the ref as attribution
      var comboBody = '';
      if (beforeRef) comboBody = beforeRef;
      if (afterRef) comboBody = comboBody ? comboBody + '\n' + afterRef : afterRef;
      if (comboBody) {
        return { body: comboBody, attribution: cln };
      }
    }
  }

  // 3. Last line starts with dash: "- Author" or "— Author"
  var lines = s.split('\n');
  var last = lines[lines.length - 1].trim();
  if (/^[-—–]\s/.test(last)) {
    return {
      body: lines.slice(0, -1).join('\n'),
      attribution: last.replace(/^[-—–\s]+/, '')
    };
  }

  // 4. Last line is a standalone name (capitalized, not a common sentence start)
  //    But NOT if it looks like a scripture reference — those should stay inline
  //    And NOT if it contains definition/content markers like = or ; or em dashes
  //    And NOT if the previous line doesn't end with punctuation (means it's a continuation)
  var commonStarts = /^(The|Our|God|He|She|It|We|You|They|This|That|For|And|But|Or|In|On|At|To|Is|As|If|So|Do|No|A|All|Every|When|What|How|Why|Not|His|Her|Its|My|Your)\s/;
  var isScriptureRef = new RegExp('^[123]?\\s*(?:' + bookNames + ')\\.?\\s+\\d', 'i');
  var isContentLine = /[=;\u2014—–]/.test(last);
  var prevLine = lines.length > 1 ? lines[lines.length - 2].trim() : '';
  var prevEndsClean = /[.!?;:)\u201D""'\u2019]$/.test(prevLine);
  if (lines.length > 1 && prevEndsClean && /^[A-Z]/.test(last) && !commonStarts.test(last) && !isScriptureRef.test(last) && !isContentLine && last.length < 60) {
    return {
      body: lines.slice(0, -1).join('\n'),
      attribution: last.replace(/^[-—–\s]+/, '')
    };
  }

  return { body: s, attribution: '' };
}
