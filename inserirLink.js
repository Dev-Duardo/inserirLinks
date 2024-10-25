function inserirLinks() {
  const folderName = "CONTRATOS 2024.03";  // Nome da pasta principal
  const abaNames = [
    "RJO009T1", "RJO022T1", "RJO066T1", "RJO084T1", "RJO099T1", 
    "RJO115T1", "RJO116B1", "ROT003T1", "RJO132T1", "RJO135T1", 
    "RJO138T1", "RJO143T1", "RJO145T1", "RJO147T1", "RJO150T1", 
    "RJO160T1", "SJO005T2", "TPO004T1", "VTR005T1", "RJO288B1", 
    "VTR024T2", "VTR034T1"
  ];

  const folder = DriveApp.getFoldersByName(folderName).next();
  const subfolders = folder.getFolders();  // Obtém as subpastas dentro da pasta principal

  // Mapeia o nome das pastas para seus links
  const folderLinks = {};
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    folderLinks[subfolder.getName()] = subfolder.getUrl();
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  abaNames.forEach((abaName) => {
    const sheet = spreadsheet.getSheetByName(abaName);
    if (sheet && folderLinks[abaName]) {
      // Cria o texto com link embutido
      const richText = SpreadsheetApp.newRichTextValue()
        .setText("CONTRATOS")
        .setLinkUrl(folderLinks[abaName])
        .build();
        
      sheet.getRange("C5").setRichTextValue(richText);
    }
  });
}
