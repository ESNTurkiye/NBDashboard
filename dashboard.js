/**
 * This function handles the click event on cells in the Google Sheets.
 * It retrieves data from the clicked cell and performs various actions based on the cell's content.
 */
function handleCellClick() {
    // Get the values of clicked cell's row and column in A1 notation.
    let sheet = SpreadsheetApp.getActiveSheet();
    let currentCell = sheet.getCurrentCell();
    let value = currentCell.getValue();
    let cellName = currentCell.getA1Notation();
    let cell = cellName[0];
    let number = parseInt(cellName.slice(1));
    let address = sheet.getRange('J' + number).getValue();
    let description = sheet.getRange('C' + number).getValue();
    let assignor = sheet.getRange('B' + number).getValue();
    let deadline = sheet.getRange('L' + number).getValue();
    let status = sheet.getRange('L' + number).getValue();
    let priority = sheet.getRange('M' + number).getValue();

    // Map of email addresses
    let emailMap = {
        "Gökay Bağrıyanık": "president@esnturkey.org",
        "Merve Ceylan": "projectmanager@esnturkey.org",
        "Gözde Özel": "nr@esnturkey.org",
        "Nisa Gökyıldız": "communication@esnturkey.org",
        "Doğukan Berk Demirdelen": "treasurer@esnturkey.org",
        "Kaan Can Yıldırım": "vicepresident@esnturkey.org",
        "Furkan Uçar": "wpa@esnturkey.org",
        "Board": "board@esnturkey.org"
    }
  
    // If the clicked cell is in column N and its value is true, send an email to the assignee.
    if (typeof value === 'boolean' && value === true) {
      // Set the body of the email to be sent.
      let body = `Assignor: ${assignor}\nTask: ${description}\nDeadline: ${deadline}`;

      // Generate a random code for the email subject. This is done to prevent duplicate email subjects, which will cause the emails to be grouped together.
      let code = Math.floor(Math.random() * 1000000)
  
      // N column is the Send Email column.
      if(cell == 'N') {
        let subject = `A new task has been assigned to you! ${code}`;
        let email = emailMap[address];
        sendEmail(email, address, subject, body, assignor);
      } 
      if (cell == 'O') {
        let email = 'auditors@esnturkey.org';
        let subject = `A task assign was made in the Board! ${code}`;
        let body = `Dear Board of Auditors,\n A task assign has been made on Dashboard. \nTask: ${description}\nAssignor: ${assignor} \nAssignee: ${address}\nDeadline:${deadline}.\n`
        MailApp.sendEmail(email, subject, body, {
          cc: getCc(assignor)
        });
  
      }
    }
  
  // If the clicked column is Status column and its value is Closed, carry the task to the Closed section of the sheet.  
  if (cell == 'L' && status == 'Closed') {
    let row = 170;
    let columnC = sheet.getRange("C" + row).getValue();
  
    while (columnC !== '') {
      row++;
      columnC = sheet.getRange("C" + row).getValue();
    }
  
    let addressCell = sheet.getRange('J' + number);
    let descriptionCell = sheet.getRange('C' + number);
    let assignorCell = sheet.getRange('B' + number);
    let deadlineCell = sheet.getRange('L' + number);
    let startdateCell = sheet.getRange('K' + number);
    let statusCell = sheet.getRange('M' + number);
  
    let addressValidation = addressCell.getDataValidation();
    let descriptionValidation = descriptionCell.getDataValidation();
    let assignorValidation = assignorCell.getDataValidation();
    let deadlineValidation = deadlineCell.getDataValidation();
    let startdateValidation = startdateCell.getDataValidation();
    let statusValidation = statusCell.getDataValidation();
  
    let targetAddressCell = sheet.getRange('J' + row);
    let targetDescriptionCell = sheet.getRange('C' + row);
    let targetAssignorCell = sheet.getRange('B' + row);
    let targetDeadlineCell = sheet.getRange('L' + row);
    let targetStartdateCell = sheet.getRange('K' + row);
    let targetStatusCell = sheet.getRange('M' + row);
  
    let value = descriptionCell.getValue();
    targetDescriptionCell.setValue(value);
    sheet.getRange('C' + row + ':I' + row).merge();
  
    let addressValue = addressCell.getValue();
    targetAddressCell.setValue(addressValue);
  
    let assignorValue = assignorCell.getValue();
    targetAssignorCell.setValue(assignorValue);
  
    let deadlineValue = deadlineCell.getValue();
    targetDeadlineCell.setValue(deadlineValue);
  
    let startdateValue = startdateCell.getValue();
    targetStartdateCell.setValue(startdateValue);
  
    let statusValue = statusCell.getValue();
    targetStatusCell.setValue(statusValue);
  
    targetAddressCell.setDataValidation(addressValidation);
    targetDescriptionCell.setDataValidation(descriptionValidation);
    targetAssignorCell.setDataValidation(assignorValidation);
    targetDeadlineCell.setDataValidation(deadlineValidation);
    targetStartdateCell.setDataValidation(startdateValidation);
    targetStatusCell.setDataValidation(statusValidation);
  
    sheet.deleteRow(number);
  }
  
  // Change the color of the cell based on the priority of the task.
  if (cell == 'M' &&  priority == 'Low') {
      changeColor(sheet, number, "#2ECC71");
    }
  
  if (cell == 'M' &&  priority == 'Normal') {
      changeColor(sheet, number, "#F6E58D");
    }
  
  if (cell == 'M' &&  priority == 'High') {
      changeColor(sheet, number, "#EB4D4B");
  }
}

  function sendEmail(email, address, subject, body, assignor) {
    if(address == "Gökay Bağrıyanık") {
        MailApp.sendEmail(email, subject, body, {
          cc: getCc(assignor)
        });
    }

    if(address == "Merve Ceylan") {
        MailApp.sendEmail(email, subject, body, {
          cc: getCc(assignor)
        });
    }
      

    if(address == "Gözde Özel") {
        MailApp.sendEmail(email, subject, body, {
          cc: getCc(assignor)
        });
    }
      

    if(address == "Nisa Gökyıldız") {
        MailApp.sendEmail(email, subject, body, {
          cc: getCc(assignor)
        });
    }

    if(address == "Doğukan Berk Demirdelen") {
        MailApp.sendEmail(email, subject, body, {
          cc: getCc(assignor)
        });
    }

    if(address == "Kaan Can Yıldırım") {
        MailApp.sendEmail(email, subject, body, {
          cc: getCc(assignor)
        });
    }

    if(address == "Furkan Uçar") {
        MailApp.sendEmail(email, subject, body, {
          cc: getCc(assignor)
        });
    }

    if(address == "Board") {
        let email = "board@esnturkey.org";
        MailApp.sendEmail(email, subject, body, {
          cc: getCc(assignor)
        });
    } 
 }

function changeColor(sheet, number, color) {
    var b = sheet.getRange('B' + number);
    b.setBackground(color);

    var c = sheet.getRange('C' + number);
    c.setBackground(color);

    var j = sheet.getRange('J' + number);
    j.setBackground(color);

    var k = sheet.getRange('K' + number);
    k.setBackground(color);

    var l = sheet.getRange('L' + number);
    l.setBackground(color);

    var m = sheet.getRange('M' + number);
    m.setBackground(color);

    var n = sheet.getRange('N' + number);
    n.setBackground(color);

    var o = sheet.getRange('O' + number);
    o.setBackground(color);

    var p = sheet.getRange('P' + number);
    p.setBackground(color);

    var q = sheet.getRange('Q' + number);
    q.setBackground(color);
}
  
  function oneDayReminder() {
    let sheet = SpreadsheetApp.getActiveSheet();
    let currentDate = new Date();
    let code = Math.floor(Math.random() * 1000000)

    let emailMap = {
      "Gökay Bağrıyanık": "president@esnturkey.org",
      "Merve Ceylan": "projectmanager@esnturkey.org",
      "Gözde Özel": "nr@esnturkey.org",
      "Nisa Gökyıldız": "communication@esnturkey.org",
      "Doğukan Berk Demirdelen": "treasurer@esnturkey.org",
      "Kaan Can Yıldırım": "vicepresident@esnturkey.org",
      "Furkan Uçar": "wpa@esnturkey.org",
      "Board": "board@esnturkey.org"
    }

    for (i=5; i < 80; i++) {
      let deadline = sheet.getRange('M' + i).getValue();
      if (deadline == currentDate && sheet.getRange('P' + i).getValue() === true) {
        let address = sheet.getRange('J' + i).getValue();
        let assignor = sheet.getRange('B' + number).getValue();
        let body = `Assignor: ${assignor}\nTask: ${description}\nDeadline: ${deadline}`;
        let subject = `Last day reminder for your task! ${code}`;

        let email = emailMap[address];
        sendEmail(email, address, subject, body, assignor);
      }
    }
  }
  
  function threeDayReminder() {
    // Milliseconds in a day
    const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    let sheet = SpreadsheetApp.getActiveSheet();
    let currentDate = new Date();
    let code = Math.floor(Math.random() * 1000000)

    let emailMap = {
      "Gökay Bağrıyanık": "president@esnturkey.org",
      "Merve Ceylan": "projectmanager@esnturkey.org",
      "Gözde Özel": "nr@esnturkey.org",
      "Nisa Gökyıldız": "communication@esnturkey.org",
      "Doğukan Berk Demirdelen": "treasurer@esnturkey.org",
      "Kaan Can Yıldırım": "vicepresident@esnturkey.org",
      "Furkan Uçar": "wpa@esnturkey.org",
      "Board": "board@esnturkey.org"
    }

    for (i=5; i < 80; i++) {
      let deadline = sheet.getRange('M' + i).getValue();
      if (deadline == currentDate - MILLIS_PER_DAY*2 && sheet.getRange('Q' + i).getValue() === true) {
        let address = sheet.getRange('J' + i).getValue();
        let body = `Assignor: ${assignor}\nTask: ${description}\nDeadline: ${deadline}`;
        let subject = `Last 3 day reminder for your task ${code}`;
        
        let email = emailMap[address];
        sendEmail(email, address, subject, body, assignor);
      }
    }
  }
  
  function getCc(assignor) {
    let email;
  
    if (assignor == "Gökay Bağrıyanık") {
      email = "president@esnturkey.org";
    } else if (assignor == "Merve Ceylan") {
      email = "projectmanager@esnturkey.org";
    } else if (assignor == "Gözde Özel") {
      email = "nr@esnturkey.org";
    } else if (assignor == "Nisa Gökyıldız") {
      email = "communication@esnturkey.org";
    } else if (assignor == "Doğukan Berk Demirdelen") {
      email = "treasurer@esnturkey.org";
    } else if (assignor == "Kaan Can Yıldırım") {
      email = "vicepresident@esnturkey.org";
    } else if (assignor == "Furkan Uçar") {
      email = "wpa@esnturkey.org";
    } else {
      email = "";
    }
  
    return email;
  }