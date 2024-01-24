await Outlook.run(async (context) => {
  const item = context.mailbox.item;
  item.body.set("Hello, world!", { coercionType: Office.CoercionType.Text });
  item.subject.set("Hello, world!");
  item.saveAsync();
  await context.sync();
});
