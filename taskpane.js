Office.onReady(function(info) {
    if (info.host === Office.HostType.Word) {
        document.getElementById("addPageNumbers").onclick = addPageNumbers;
    }
});

function addPageNumbers() {
    Word.run(function(context) {
        // Get the document sections
        var sections = context.document.sections;
        var body = context.document.body;
        sections.load('items');

        return context.sync().then(function() {
            if (sections.items.length > 3) { // Ensure the document has at least 3 sections
                var section = sections.items[3]; // Access the third section
                var footer = section.getFooter("Primary");

                // Clear existing footer content
                footer.clear();

                // Insert page numbers in the format "Page X of 79"
                for (var i = 1; i <= 79; i++) {
                    // Insert "Page X of 79" text
                    footer.insertText("Page " + i + " of 79", Word.InsertLocation.end);

                    // If it's not the last page, insert a page break to simulate the next page
                    if (i < 79) {
                        body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
                    }
                }

                return context.sync();
            } else {
                console.log("The document does not have a third section.");
            }
        });
    }).catch(function(error) {
        console.log("Error: " + JSON.stringify(error, null, 2));
        console.log("Error message: " + error.message);
        console.log("Stack trace: " + error.stack);
    });
}

