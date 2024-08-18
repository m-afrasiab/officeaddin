Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
        document.getElementById("run").onclick = run;
    }
});

async function run() {
    await Word.run(async (context) => {
        // Get all sections in the document
        const sections = context.document.sections;
        sections.load('items');

        await context.sync();

        // Get the last section
        const lastSection = sections.items[sections.items.length - 1];

        // Get all paragraphs in the last section and load necessary properties
        const paragraphs = lastSection.body.paragraphs;
        paragraphs.load('items/style');

        // Load the listItem property to check if a paragraph is part of a list
        paragraphs.items.forEach(paragraph => {
            paragraph.load('listItem');
        });

        await context.sync();

        // Iterate through the paragraphs and add an empty line after each, excluding headings and lists
        paragraphs.items.forEach(paragraph => {
            const style = paragraph.style;

            // Skip headings (styles start with "Heading") and list paragraphs
            if (!style.startsWith("Heading") && !paragraph.listItem) {
                paragraph.insertParagraph("", Word.InsertLocation.after);
            }
        });

        await context.sync();
    }).catch(function (error) {
        console.log("Error: " + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
