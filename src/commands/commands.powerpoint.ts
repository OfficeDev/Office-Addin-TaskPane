await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getFirst();
    slide.shapes.getFirst().text = "Hello, world!";
    await context.sync();
});