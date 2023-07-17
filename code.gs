// Get the Id from the URL
// font: https://stackoverflow.com/questions/16840038/easiest-way-to-get-file-id-from-url-on-google-apps-script
function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }

//Principal Function
function createOneSlidePerRow() {

  // Load data from the spreadsheet.
  let dataRange = SpreadsheetApp.getActive().getDataRange();
  let sheetContents = dataRange.getValues();

  // Save the header in a variable called header
  let header = sheetContents.shift();

  // Create an array to save the data to be written back to the sheet.
  // We'll use this array to save links to the slides that are created.
  let updatedContents = [];

  // Reverse the order of rows because new slides will
  // be inserted at the top. Without this, the order of slides
  // will be the inverse of the ordering of rows in the sheet. 
  sheetContents.reverse();

  // For every row, create a new slide by duplicating the master slide
  // and replace the template variables with data from that row.
  sheetContents.forEach(function (row) {
  // Get the Master ID from the slides
  let masterDeckURL = row[3];
  masterDeckID = getIdFromUrl(masterDeckURL);

  // Get the Master ID from the folder
  let folderURL = row[4];
  folderId = getIdFromUrl(folderURL);

  // Open the presentation and get the slides in it.
  let deck = SlidesApp.openById(masterDeckID);
  let slides = deck.getSlides();

  // The 2nd slide is the template that will be duplicated
  // once per row in the spreadsheet.
  slides[0].duplicate;
  

  // 2. Create a temporal Google Slides.
  let file = DriveApp.getFileById(masterDeckID).makeCopy("temp");
  let id = file.getId();
  let temp = SlidesApp.openById(id);
  let tempSlides = temp.getSlides();
  

    // Populate data in the slide that was created
    temp.replaceAllText("{{Text1}}", row[0]);
    temp.replaceAllText("{{Text2}}", row[1]);
    temp.replaceAllText("{{Text3}}", row[2])

tempSlides[1].remove()

    // 3. Export each page as a PDF file.
  temp.saveAndClose();
  let blob = file.getBlob().setName(`${row[0]+"_"+row[2]}.pdf`);

  // 4. Remove the temporal Google Slides.
  file.setTrashed(true);

  // Create the URL for the slide using the deck's ID and the ID
  // of the slide.
  
  let pdfUrl = DriveApp.getFolderById(folderId).createFile(blob).getUrl();

  // Add this URL to the 4th column of the row and add this row
  // to the data to be written back to the sheet.
  row[5] = pdfUrl;
  updatedContents.push(row);

  slides[1].remove;

});

  // Add the header back (remember it was removed using 
  // sheetContents.shift())
  updatedContents.push(header);

  // Reverse the array to preserve the original ordering of 
  // rows in the sheet.
  updatedContents.reverse();

  // Write the updated data back to the Google Sheets spreadsheet.
  dataRange.setValues(updatedContents);

}
