await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.font.color = "red";
    await context.sync();
});