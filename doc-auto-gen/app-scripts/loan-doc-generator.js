function onOpen() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('Generate Doc')
  menu.addItem('Create New Doc', 'createNewGoogleDoc')
  menu.addToUi()
}

function createNewGoogleDoc() {
  //This value should be the id of your document template that we created in the last step
  const googleDocTemplate = DriveApp.getFileById(
    '1Ww5EPC_qnvYo4ZsMWfcfxS5aSsrCa1AgwyAyqUX38mU'
  )

  //This value should be the id of the folder where you want your completed documents stored
  const destinationFolder = DriveApp.getFolderById(
    '1CtplWdU9yVRqjPRDMwaixOiKZ6WY00b7'
  )
  const data = {}
  //Here we store the sheet as a variable
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Loan info')
  const loanCalculationSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Details and table')

  // process borrower and cosigner information
  const borrowerInfoRows = sheet.getRange(11, 1, 5, 8).getValues()
  for (let i = 1; i < borrowerInfoRows.length; i++) {
    for (let j = 1; j < borrowerInfoRows[i].length; j++) {
      const name = `${borrowerInfoRows[i][0]} ${borrowerInfoRows[0][j]}`
      data[name.trim()] = borrowerInfoRows[i][j]
    }
  }

  // process loan calculations
  const loanCalculationRows = loanCalculationSheet
    .getRange(1, 1, 30, 2)
    .getValues()
  Logger.log(loanCalculationRows)
  loanCalculationRows.forEach((r) => {
    if (r[0]) {
      const key = r[0].trim()
      if (typeof r[1]?.getMonth === 'function') {
        data[key] = r[1].toISOString().split('T')[0]
      } else if (typeof r[1] == 'number' && !r[0].includes('number')) {
        data[key] = (Math.round(r[1] * 100) / 100).toFixed(2)
      } else {
        data[key] = r[1]
      }
    }
  })

  Object.keys(data).forEach((k) => Logger.log(`${k}: ${data[k]}`))

  const date = new Date()

  //Using the row data in a template literal, we make a copy of our template document in our destinationFolder
  const copy = googleDocTemplate.makeCopy(
    `${data['Loan number']}-contract-${date}`,
    destinationFolder
  )
  //Once we have the copy, we then open it using the DocumentApp
  const doc = DocumentApp.openById(copy.getId())
  //All of the content lives in the body, so we get that for editing
  const body = doc.getBody()

  //In these lines, we replace our replacement tokens with values from our spreadsheet row
  for (let key in data) {
    if (key && data[key]) {
      body.replaceText(`{{${key}}}`, data[key])
    }
  }

  //We make our changes permanent by saving and closing the document
  doc.saveAndClose()
  //Store the url of our new document in a variable
  const url = doc.getUrl()
  //Write that value back to the 'Document Link' column in the spreadsheet.
  sheet.getRange(20, 1).setValue(url)
}
