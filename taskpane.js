Office.onReady(function(info) {
    if (info.host === Office.HostType.Word) {
        document.getElementById("run").onclick = run;
    }
});

async function run() {
    await Word.run(async (context) => {
        // Get all paragraphs in the document
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        
        await context.sync();

        // Insert an empty line after each paragraph
        paragraphs.items.forEach(paragraph => {
            paragraph.insertParagraph("", Word.InsertLocation.after);
        });

        await context.sync();
    });
}
