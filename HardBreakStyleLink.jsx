#target indesign

function main() {
    var credit = "Made with \u2764\uFE0F by Rajesh Budhathoki";

    function showMessage(msg, isError) {
        var prefix = isError ? "Error: " : "Success: ";
        alert(prefix + msg + "\n\n" + credit);
    }

    try {
        if (app.selection.length === 0 || !app.selection[0].hasOwnProperty("baseline")) {
            showMessage("Please select some text or a text frame first.", true);
            return;
        }

        var doc = app.activeDocument;
        var selection = app.selection[0];
        var targetText = (selection.constructor.name === "TextFrame") ? selection.texts[0] : selection;

        app.doScript(function() {
            try {
                targetText.convertBulletsAndNumberingToText();
            } catch(e) {}

            var ranges = targetText.textStyleRanges;
            for (var i = 0; i < ranges.length; i++) {
                var cStyle = ranges[i].appliedCharacterStyle;
                if (cStyle.name !== "[None]") {
                    ranges[i].appliedCharacterStyle = cStyle;
                }
            }

            targetText.applyParagraphStyle(doc.paragraphStyles.item(0), false);

        }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Hard Break Link to Style");

        showMessage("Style links broken successfully while preserving appearance.", false);

    } catch (err) {
        showMessage(err.message, true);
    }
}

main();