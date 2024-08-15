Office.onReady(function(info) {
    if (info.host === Office.HostType.Word) {
        document.getElementById("addPageNumbers").onclick = addPageNumbers;
    }
});

function addPageNumbers() {
    Word.run(function(context) {
        // Load the paragraphs in the document
        var paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');

        return context.sync().then(function() {
            var pageNumber = 1;
            var previousParagraph;

            paragraphs.items.forEach(function(paragraph, index) {
                // Load the paragraph's range to find its position
                var range = paragraph.getRange();
                range.load('text');

                return context.sync().then(function() {
                    // Insert page number at the end of the paragraph if it's at the end of the page
                    if (isEndOfPage(paragraph, previousParagraph)) {
                        paragraph.insertText(" - Page " + pageNumber, Word.InsertLocation.end);
                        pageNumber++;
                    }

                    previousParagraph = paragraph;
                });
            });

            return context.sync();
        });
    }).catch(function(error) {
        console.log("Error: " + JSON.stringify(error, null, 2));
        console.log("Error message: " + error.message);
        console.log("Stack trace: " + error.stack);
    });
}

function isEndOfPage(currentParagraph, previousParagraph) {
    // Placeholder logic to determine if the paragraph is at the end of a page.
    // This can be customized based on the document's layout.
    
    // Simple heuristic: If the current paragraph is in a new section or is far from the previous paragraph
    // (which would usually indicate a page break or a large space), we consider it the end of the page.
    
    if (!previousParagraph) return false;

    // Calculate the difference in paragraph positions to guess a page break
    var previousRange = previousParagraph.getRange();
    var currentRange = currentParagraph.getRange();

    return (currentRange.text !== previousRange.text);
}
