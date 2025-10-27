// The initialize function must be run each time a new page is loaded.
Office.onReady(() => {
        // If you need to initialize something you can do so here.
});

async function sampleFunction(event) {
try {
        await Word.run(async (context) => {
        // Insert a paragraph at the end of the document body.
        const body = context.document.body;
        body.insertParagraph("You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents.", Word.InsertLocation.end);

        await context.sync();
        });
} catch (error) {
        console.error(error);
}

// Calling event.completed is required. event.completed lets the platform know that processing has completed.
event.completed();
}
