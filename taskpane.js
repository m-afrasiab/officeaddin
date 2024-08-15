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

                // Insert "Page " text
                var pageRange = footer.insertParagraph("Page ", Word.InsertLocation.end);

                // Insert page number field
                var pageField = pageRange.insertField(Word.FieldType.page, true);

                // Insert " of " text
                footer.insertText(" of ", Word.InsertLocation.end);

                // Insert total number of pages field
                var totalPagesField = footer.insertField(Word.FieldType.numPages, true);

                // Set the starting page number for the second section to 1
                section.pageSetup.startingNumber = 1;

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
