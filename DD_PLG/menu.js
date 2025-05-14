function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🔌PLG list')
    .addItem('GitHub Plugin Branches', 'getAllBranches')
    .addToUi();

    ui.createMenu('🧠AI解析')
    .addItem('プラグ解析くん', 'analyzePlg')
    .addToUi(); 
}
