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
            if (sections.items.length > 0) {
                var totalPages = 79; // Total number of pages
                var currentPage = 1;

                // Loop through each section
                sections.items.forEach(function(section) {
                    var footer = section.getFooter("Primary");

                    // Loop through each page in the section and add the page number
                    for (var i = 0; i < totalPages && currentPage <= totalPages; i++) {
                        // Insert "Page X of 79" text
                        footer.insertText("Page " + currentPage + " of " + totalPages, Word.InsertLocation.end);

                        // Move to the next page
                        currentPage++;

                        // Insert a section break to simulate moving to the next page (optional)
                        if (currentPage <= totalPages) {
                            section.insertBreak(Word.BreakType.sectionNextPage, Word.InsertLocation.end);
                        }
                    }
                });

                return context.sync();
            } else {
                console.log("The document does not have any sections.");
            }
        });
    }).catch(function(error) {
        console.log("Error: " + JSON.stringify(error, null, 2));
        console.log("Error message: " + error.message);
        console.log("Stack trace: " + error.stack);
    });
}
