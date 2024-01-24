await Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  range.format.fill.color = "yellow";
  await context.sync();
});
