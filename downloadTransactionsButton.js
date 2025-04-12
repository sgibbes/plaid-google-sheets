  // // download transactions
  // if (editedSheet.getName() === 'runScript' && editedCell === 'TRUE') {
  //     SpreadsheetApp.getActiveSpreadsheet().toast('stuff')

  //   editedSheet.getRange(1, 2).setValue("FALSE");
  //   const monthToDownload = editedSheet.getRange(2,2).getValue()

  //   if (monthToDownload) {
  //     const month = monthToDownload.split('-')[0]
  //     const year = monthToDownload.split('-')[1]
  //     getRealTransactions(month, year)

  //   } else {
  //     getRealTransactions()
  //   }

  //   isRunning = false;
  //   return;

  // }