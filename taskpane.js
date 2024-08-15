Office.onReady(function(info) {
    if (info.host === Office.HostType.Word) {
        document.getElementById("addPageNumbers").onclick = addPageNumbers;
    }
});

function addPageNumbers() {
    Word.run(function(context) {
        // Get the document body and sections
        var sections = context.document.sections;
        sections.load('items');

        return context.sync().then(function() {
            if (sections.items.length > 1) {
                var section = sections.items[3]; // Get the second section
                var footer = section.getFooter("Primary");

                // Loop to add page numbers from 1 to 79
                for (var i = 1; i <= 79; i++) {
                    // Insert "Page X of 79" text
                    footer.insertText("Page " + i + " of 79", Word.InsertLocation.end);
                    
                    // Insert a line break after each page number (optional, depending on layout)
                    footer.insertBreak(Word.BreakType.line, Word.InsertLocation.end);
                }

                // Set the starting page number for the second section to 1


                return context.sync();
            } else {
                console.log("The document does not have enough pages to split.");
            }
        });
    }).catch(function(error) {
        console.log("Error: " + JSON.stringify(error, null, 2));
        console.log("Error message: " + error.message);
        console.log("Stack trace: " + error.stack);
    });
}
