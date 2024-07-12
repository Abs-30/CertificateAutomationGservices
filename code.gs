function generateCertificates() {
  // Replace placeholders with actual IDs
  var slideId = 'Slide ID'; // ID of the Slide template
  var spreadsheetId = 'Sheet ID'; // ID of the Spreadsheet
  var destinationFolderId = 'Folder ID'; // ID of the destination folder for PDF certificates

  // Accessing necessary resources
  var destinationFolder = DriveApp.getFolderById(destinationFolderId);
  var slide = DriveApp.getFileById(slideId);
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheets()[0]; // Assuming the data is on the first sheet

  // Fetch data from the spreadsheet
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Remove header row
  data.shift();

  // Loop through each row of data
  data.forEach(function (row) {
    var name = row[0];
    var email = row[1];
    var team = row[2]; 

    // Make a copy of the template slide
    var newSlide = slide.makeCopy(name + ' Certificate', destinationFolder);

    // Access the copied slide and replace placeholders with actual data
    var presentation = SlidesApp.openById(newSlide.getId());
    var slides = presentation.getSlides();
    var placeholders = {
      "{{name}}": name,
      "{{team}}": team
    };

    // Iterate through slide shapes to find and replace placeholder text
    slides.forEach(function (slide) {
      var shapes = slide.getShapes();
      shapes.forEach(function (shape) {
        try {
          if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
            var texts = shape.getText();
            Object.keys(placeholders).forEach(function (placeholder) {
              if (texts.asString().includes(placeholder)) {
                var newText = texts.asString().replace(placeholder, placeholders[placeholder]);
                shape.getText().setText(newText);
              }
            });
          }
        } catch (error) {
          console.error("Error processing shape: " + shape.getObjectId(), error);
        }
      });
    });

    // Save changes to the modified slide
    presentation.saveAndClose();
    presentation = SlidesApp.openById(newSlide.getId());

    // Export the modified slide as PDF
    var pdf = DriveApp.getFileById(newSlide.getId()).getAs('application/pdf');
    destinationFolder.createFile(pdf);

    // Email the certificate to the recipient, handling invalid email address
    try {
      MailApp.sendEmail({
        to: email,
        subject: 'Congratulations on Your Outstanding Achievement',
        body: `Dear ${name},\n\nCongrats on your achievement with Team ${team}!\n\nBest regards,\nYour Organization`,
        attachments: [pdf]
      });
      console.log(`Sent certificate to ${name} (${email})`);
    } catch (error) {
      console.error(`Error sending email to ${name} (${email}):`, error);
    }
  });
}
