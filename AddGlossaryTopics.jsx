// AddGlossaryTopics.jsx

function main() {
  if (app.documents.length === 0) {
    alert("Please open an InDesign document first.");
    return;
  }

  var doc = app.activeDocument;

  var allStyles = doc.allParagraphStyles;
  var styleNames = ["[Any Style]"];
  for (var s = 0; s < allStyles.length; s++) {
    styleNames.push(allStyles[s].name);
  }

  var dialog = new Window("dialog", "Add Index Topics & Page Refs");
  dialog.orientation = "column";
  dialog.alignChildren = ["fill", "top"];

  var stylePanel = dialog.add(
    "panel",
    undefined,
    "Search Filters (Leave blank/default to ignore)",
  );
  stylePanel.orientation = "column";
  stylePanel.alignChildren = ["left", "center"];
  stylePanel.margins = 15;

  // Paragraph Style Dropdown
  var styleGroup = stylePanel.add("group");
  styleGroup.add("statictext", undefined, "Target Paragraph Style:");
  var styleDropdown = styleGroup.add("dropdownlist", undefined, styleNames);
  styleDropdown.selection = 0;

  // Page Range Inputs
  var pageGroup = stylePanel.add("group");
  pageGroup.add("statictext", undefined, "Start Page:");
  var startPageInput = pageGroup.add("edittext", undefined, doc.pages[0].name);
  startPageInput.characters = 5;

  pageGroup.add("statictext", undefined, " End Page:");
  var endPageInput = pageGroup.add(
    "edittext",
    undefined,
    doc.pages[doc.pages.length - 1].name,
  );
  endPageInput.characters = 5;

  // Match Options Checkboxes
  var optionsGroup = stylePanel.add("group");
  optionsGroup.alignment = ["left", "center"];
  var chkCaseSensitive = optionsGroup.add(
    "checkbox",
    undefined,
    "Case Sensitive",
  );
  chkCaseSensitive.value = true; // Checked by default

  var chkWholeWord = optionsGroup.add("checkbox", undefined, "Whole Word");
  chkWholeWord.value = true; // Checked by default

  var dataPanel = dialog.add("panel", undefined, "Index Data");
  dataPanel.orientation = "column";
  dataPanel.alignChildren = ["fill", "top"];
  dataPanel.margins = 15;
  dataPanel.add(
    "statictext",
    undefined,
    "Paste your index data below (Format: Word | Meaning):",
  );

  var dataInput = dataPanel.add("edittext", undefined, "", {
    multiline: true,
    scrolling: true,
  });
  dataInput.preferredSize = [400, 200];

  var buttonRow = dialog.add("group");
  buttonRow.alignment = "fill";
  var leftGroup = buttonRow.add("group");
  leftGroup.alignment = ["left", "center"];
  var btnClear = leftGroup.add("button", undefined, "Delete All & Start Fresh");
  var rightGroup = buttonRow.add("group");
  rightGroup.alignment = ["right", "center"];
  var btnCancel = rightGroup.add("button", undefined, "Cancel", {
    name: "cancel",
  });
  var btnRun = rightGroup.add("button", undefined, "Run Indexer", {
    name: "ok",
  });

  // UI Instructions added here
  var instructionsText = dialog.add(
    "statictext",
    undefined,
    "Instructions:\n1. Select a Target Style, Page Range, and Search Options.\n2. Paste your data in 'Word | Meaning' format (one per line).\n3. Click 'Run Indexer' to automatically create markers.",
    { multiline: true },
  );
  instructionsText.alignment = ["fill", "center"];

  var creditText = dialog.add(
    "statictext",
    undefined,
    "Made with \u2764\uFE0F by Rajesh Budhathoki",
  );
  creditText.alignment = "center";

  btnClear.onClick = function () {
    if (
      confirm(
        "WARNING: This will DELETE ALL previous index topics, page references, and symbols from the document and Index panel.\n\nThis action CANNOT BE UNDONE.\n\nContinue and start fresh?",
      )
    ) {
      clearAllIndexData();
      alert(
        "All previous symbols and indexes removed. Index panel should now be empty. You can add fresh data.",
      );
    }
  };

  if (dialog.show() === 1) {
    var rawText = dataInput.text;

    if (rawText.replace(/\s/g, "") === "") {
      alert("No data provided. Script cancelled.");
      return;
    }

    var selectedStyleIndex = styleDropdown.selection.index;
    var selectedStyle =
      selectedStyleIndex === 0 ? null : allStyles[selectedStyleIndex - 1];

    var startPageName =
      startPageInput.text !== "" ? startPageInput.text : doc.pages[0].name;
    var endPageName =
      endPageInput.text !== ""
        ? endPageInput.text
        : doc.pages[doc.pages.length - 1].name;

    // Grab checkbox values
    var isCaseSensitive = chkCaseSensitive.value;
    var isWholeWord = chkWholeWord.value;

    var rawArray = rawText.split("\n");
    var indexData = [];

    for (var k = 0; k < rawArray.length; k++) {
      var line = rawArray[k].replace(/^\s+|\s+$/g, "");

      if (line !== "" && line.indexOf("|") !== -1) {
        var parts = line.split("|");
        var cleanWord = parts[0].replace(/^\s+|\s+$/g, "");
        var cleanMeaning = parts[1].replace(/^\s+|\s+$/g, "");

        if (cleanWord !== "" && cleanMeaning !== "") {
          var formattedLine = cleanWord + " | " + cleanMeaning;
          indexData.push({ fullTopic: formattedLine, searchWord: cleanWord });
        }
      }
    }

    if (indexData.length === 0) {
      alert(
        "No valid data found. Ensure you are using the 'Word | Meaning' format.",
      );
      return;
    }

    app.doScript(
      function () {
        runCoreIndexer(
          indexData,
          selectedStyle,
          startPageName,
          endPageName,
          isCaseSensitive,
          isWholeWord,
        );
      },
      ScriptLanguage.JAVASCRIPT,
      undefined,
      UndoModes.ENTIRE_SCRIPT,
      "Auto Index with Page Refs",
    );
  }
}

// Upgraded getPage function: Safely punches through groups and tables
function getPage(obj) {
  if (
    obj.hasOwnProperty("parentTextFrames") &&
    obj.parentTextFrames.length > 0
  ) {
    var tf = obj.parentTextFrames[0];
    if (tf.hasOwnProperty("parentPage") && tf.parentPage !== null) {
      return tf.parentPage;
    }
  }

  while (obj) {
    if (obj instanceof Page) return obj;
    if (obj.hasOwnProperty("parentPage") && obj.parentPage !== null)
      return obj.parentPage;
    if (obj instanceof Application) return null;
    obj = obj.parent;
  }
  return null;
}

function runCoreIndexer(
  indexData,
  targetStyle,
  startPageName,
  endPageName,
  isCaseSensitive,
  isWholeWord,
) {
  var doc = app.activeDocument;
  var myIndex = doc.indexes.length === 0 ? doc.indexes.add() : doc.indexes[0];

  var startOffset = 0;
  var endOffset = doc.pages.length - 1;

  for (var p = 0; p < doc.pages.length; p++) {
    if (doc.pages[p].name === startPageName)
      startOffset = doc.pages[p].documentOffset;
    if (doc.pages[p].name === endPageName)
      endOffset = doc.pages[p].documentOffset;
  }

  if (startOffset > endOffset) {
    var temp = startOffset;
    startOffset = endOffset;
    endOffset = temp;
  }

  // Clear preferences to ensure native Find/Replace acts cleanly
  app.findTextPreferences = NothingEnum.nothing;
  app.changeTextPreferences = NothingEnum.nothing;

  // Apply User UI Selections
  app.findChangeTextOptions.caseSensitive = isCaseSensitive;
  app.findChangeTextOptions.wholeWord = isWholeWord;

  if (targetStyle !== null) {
    app.findTextPreferences.appliedParagraphStyle = targetStyle;
  }

  var newTopicsAdded = 0;
  var existingTopicsFound = 0;
  var newRefsAdded = 0;

  for (var i = 0; i < indexData.length; i++) {
    var currentTopicString = indexData[i].fullTopic;
    var wordToSearch = indexData[i].searchWord;
    var myTopic = null;

    myTopic = myIndex.topics.itemByName(currentTopicString);

    if (!myTopic.isValid) {
      myTopic = myIndex.topics.add(currentTopicString);
      newTopicsAdded++;
    } else {
      existingTopicsFound++;
    }

    // Tracker object to prevent duplicate references on the same page
    var processedPages = {};

    // Using native Find parameter to locate keyword
    app.findTextPreferences.findWhat = wordToSearch;
    var foundItems = doc.findText();

    for (var j = foundItems.length - 1; j >= 0; j--) {
      var foundText = foundItems[j];
      var textPage = getPage(foundText);

      if (textPage !== null) {
        var currentOffset = textPage.documentOffset;
        var pageId = textPage.id; // Get the unique ID of the current page

        if (currentOffset >= startOffset && currentOffset <= endOffset) {
          // Only process if we haven't added a reference for this topic on this page yet
          if (!processedPages[pageId]) {
            try {
              // Add reference to the exact insertion point (creates the ~I marker natively)
              myTopic.pageReferences.add(
                foundText.insertionPoints[0],
                PageReferenceType.CURRENT_PAGE,
              );
              newRefsAdded++;
              processedPages[pageId] = true; // Mark this page as processed
            } catch (e) {
              // Fallback: If insertion point fails on a ligature, apply to the word object
              try {
                myTopic.pageReferences.add(
                  foundText,
                  PageReferenceType.CURRENT_PAGE,
                );
                newRefsAdded++;
                processedPages[pageId] = true; // Mark this page as processed
              } catch (e2) {}
            }
          }
        }
      }
    }
  }

  // Force InDesign to recalculate page numbers natively
  if (doc.indexes.length > 0) {
    try {
      doc.indexes[0].update();
    } catch (e) {}
  }

  // Force a complete UI redraw to ensure the Index panel catches up
  app.scriptPreferences.enableRedraw = false;
  app.scriptPreferences.enableRedraw = true;

  app.findTextPreferences = NothingEnum.nothing;
  app.changeTextPreferences = NothingEnum.nothing;

  var styleUsedString = targetStyle !== null ? targetStyle.name : "[Any Style]";
  var matchString =
    (isCaseSensitive ? "Yes" : "No") + " / " + (isWholeWord ? "Yes" : "No");

  alert(
    "Indexing Complete!\n\n" +
      "Paragraph Style: " +
      styleUsedString +
      "\n" +
      "Page Range: " +
      startPageName +
      " to " +
      endPageName +
      "\n" +
      "Case/Word Match: " +
      matchString +
      "\n\n" +
      "New Topics Created: " +
      newTopicsAdded +
      "\n" +
      "Existing Topics Updated: " +
      existingTopicsFound +
      "\n" +
      "New Page References Added: " +
      newRefsAdded +
      "\n\n" +
      "If the new topics do not appear in the Index panel (Symbols section):\n" +
      "- Open Window > Type & Tables > Index\n" +
      "- Toggle between Reference and Topic\n" +
      "- Close and re-open the Index panel\n" +
      "- Use panel menu > Generate Index (you can cancel the dialog)",
  );
}

function clearAllIndexData() {
  var d = app.activeDocument;
  var oldRedraw = app.scriptPreferences.enableRedraw;
  app.scriptPreferences.enableRedraw = false;
  app.doScript(
    function () {
      app.findGrepPreferences = NothingEnum.nothing;
      app.changeGrepPreferences = NothingEnum.nothing;
      app.findGrepPreferences.findWhat = "~I";
      d.changeGrep();
      app.findGrepPreferences = NothingEnum.nothing;
      app.changeGrepPreferences = NothingEnum.nothing;

      if (d.indexes.length > 0) {
        try {
          d.indexes[0].topics.everyItem().remove();
        } catch (e) {}
      }
    },
    ScriptLanguage.JAVASCRIPT,
    undefined,
    UndoModes.ENTIRE_SCRIPT,
    "Delete All Index Data",
  );
  app.scriptPreferences.enableRedraw = oldRedraw;
}

main();
