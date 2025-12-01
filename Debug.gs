function debugListBookmarks() {
  const doc = DocumentApp.getActiveDocument();
  const bms = doc.getBookmarks();

  Logger.log("===== Bookmarks in this document =====");
  bms.forEach((b, i) => {
    Logger.log(`${i}: id=${b.getId()}  elementText="${b.getPosition().getElement().getText()}"`);
  });
}
