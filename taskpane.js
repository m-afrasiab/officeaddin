Office.onReady(function(info) {
    if (info.host === Office.HostType.Word) {
        document.getElementById("addPageNumbers").onclick = addPageNumbers;
    }
});

function addPageNumbers() {
    Word.run(function(context) {
        // Get the document body
        var body = context.document.body;
        context.load(body);

        // Split the document into sections at the 8th page
        return context.sync().then(function() {
            // Assuming the first 8 pages are in one section
            var sections = context.document.sections;
            sections.load('items');
            return context.sync().then(function() {
                if (sections.items.length > 1) {
                    var section = sections.items[3]; // Get the second section
                    var footer = section.getFooter("Primary");

                    // Add page number starting from this section
                    footer.insertText("Page ", Word.InsertLocation.end);
                    footer.fields.add(footer.getRange(), Word.FieldType.pageNumber);
                    footer.insertText(" of ", Word.InsertLocation.end);
                    footer.fields.add(footer.getRange(), Word.FieldType.numPages);

                    // Set the starting page number for the second section to 1
                    section.pageSetup.startingNumber = 1;

                    return context.sync();
                } else {
                    console.log("The document does not have enough pages to split.");
                }
            });
        });
    }).catch(function(error) {
        console.log("Error: " + JSON.stringify(error));
    });
}
