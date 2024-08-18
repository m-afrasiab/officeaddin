Office.onReady(function(info) {
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

        // Get all paragraphs in the last section
        const paragraphs = lastSection.body.paragraphs;
        paragraphs.load('items');

        await context.sync();

        // Insert an empty line after each paragraph in the last section
        paragraphs.items.forEach(paragraph => {
            paragraph.insertParagraph("", Word.InsertLocation.after);
        });

        await context.sync();
    })
    .catch(function(error) {
        console.log("Error: " + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
