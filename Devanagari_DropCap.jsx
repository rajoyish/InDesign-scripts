#target indesign

function createDevanagariDropCap() {
    if (app.documents.length === 0) {
        alert("Please open a document.");
        return;
    }
    if (app.selection.length === 0 || !app.selection[0].hasOwnProperty("baseline")) {
        alert("Please place your text cursor inside the paragraph you want to format.");
        return;
    }

    var doc = app.activeDocument;
    var sel = app.selection[0];
    var para = sel.paragraphs[0];

    var dialog = app.dialogs.add({name: "Devanagari Custom Drop Cap", canCancel: true});
    
    with(dialog.dialogColumns.add()){
        with(dialogRows.add()){
            staticTexts.add({staticLabel: "Characters to Drop Cap:"});
            var charInput = textEditboxes.add({editContents: "व्य", minWidth: 100});
        }
        with(dialogRows.add()){
            staticTexts.add({staticLabel: "Number of Lines to Drop:"});
            var linesInput = realEditboxes.add({editValue: 3});
        }
        with(dialogRows.add()){
            staticTexts.add({staticLabel: "Space After Drop Cap (mm):"});
            var spaceInput = realEditboxes.add({editValue: 2.5});
        }
        with(dialogRows.add()){
            var credit = "Made with \u2764\uFE0F by Rajesh Budhathoki";
            staticTexts.add({staticLabel: credit});
        }
    }

    if (dialog.show() === true) {
        var dcChars = charInput.editContents;
        var dcLines = linesInput.editValue;
        var dcSpace = spaceInput.editValue;
        dialog.destroy();

        app.doScript(function(){
            
            if (para.contents.substring(0, dcChars.length) === dcChars) {
                para.characters.itemByRange(0, dcChars.length - 1).remove();
            }

            var lineSpace = para.pointSize;
            if (para.leading !== Leading.AUTO) {
                lineSpace = para.leading;
            } else {
                lineSpace = para.pointSize * 1.2; 
            }
            var targetHeight = lineSpace * dcLines;

            var tf = para.insertionPoints[0].textFrames.add();
            tf.contents = dcChars;

            var tfPara = tf.paragraphs[0];
            tfPara.appliedFont = para.appliedFont;
            tfPara.fontStyle = para.fontStyle;
            tfPara.pointSize = targetHeight * 0.75; 

            tf.textFramePreferences.verticalJustification = VerticalJustification.TOP_ALIGN;
            tf.fit(FitOptions.FRAME_TO_CONTENT);

            tf.textWrapPreferences.textWrapMode = TextWrapModes.BOUNDING_BOX_TEXT_WRAP;
            tf.textWrapPreferences.textWrapOffset = [0, 0, 0, dcSpace + "mm"]; 

            tf.anchoredObjectSettings.anchoredPosition = AnchorPosition.ANCHORED;
            tf.anchoredObjectSettings.anchorPoint = AnchorPoint.TOP_LEFT_ANCHOR;
            
            // Fix: Separated into HorizontallyRelativeTo and VerticallyRelativeTo
            tf.anchoredObjectSettings.horizontalReferencePoint = HorizontallyRelativeTo.TEXT_FRAME;
            tf.anchoredObjectSettings.verticalReferencePoint = VerticallyRelativeTo.LINE_BASELINE;
            
            tf.anchoredObjectSettings.anchorYoffset = 0; 
            tf.anchoredObjectSettings.anchorXoffset = 0;

        }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Devanagari Drop Cap");
    } else {
        dialog.destroy();
    }
}

createDevanagariDropCap();