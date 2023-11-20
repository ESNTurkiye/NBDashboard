# NB Dashboard

Google Apps Script ile hazırlamış olduğunuz Google Sheets'e eklemeniz gerekmektedir.

Adres ve Atanan kişi bilgileri dropdown menü yapılabilir, bu sayede zaman ve kod maliyetimizi azaltmış oluruz.

```
function handleCellClick() {
    // mevcut sheet
    let sheet = SpreadsheetApp.getActiveSheet();
    // üzerinde bulunduğumuz cell
    let currentCell = sheet.getCurrentCell();
    // üzerinde bulunduğumuz cell'in içerisindeki değer
    let value = currentCell.getValue();
    // üzerinde bulunduğumuz cell'in notation adı (örn: A1)
    let cellName = currentCell.getA1Notation();
    // üzerinde bulunduğumuz cell'in sütun adı (örn: A)
    let cell = cellName[0];
    // üzerinde bulunduğumuz cell'in satır numarası (örn: 1)
    let number = parseInt(cellName.slice(1));

    /**
     * J sütununda isimleri tuttuğumuz için, seçili adresi alıyoruz.
     */
    let address = sheet.getRange('J' + number).getValue();

    /**
     * C sütununda iş açıklamasını tuttuğumuz için, seçili açıklamayı alıyoruz.
     */
    let description = sheet.getRange('C' + number).getValue();

    /**
     * B sütununda iş atayanı tuttuğumuz için, seçili atayanı alıyoruz.
     */
    let assigner = sheet.getRange('B' + number).getValue();

    /**
     * L sütununda iş deadline'ını tuttuğumuz için, seçili deadline'ı alıyoruz.
     */
    let deadline = sheet.getRange('L' + number).getValue();

    /**
     * K sütununda iş başlangıç tarihini tuttuğumuz için, seçili başlangıç tarihini alıyoruz.
     */
    let startDate = sheet.getRange('K' + number).getValue();

    /**
     * M sütununda iş durumunu tuttuğumuz için, seçili durumu alıyoruz.
     */
    let status = sheet.getRange('M' + number).getValue();
  
    /**
     * Eğer aksiyonu yaptığımız cell bir checkbox ise ve değeri boş değilse if bloğuna giriyoruz.
     */
    if (typeof value === 'boolean' && value === true) {
  
      /**
       * Üstte tanımladığımız değişkenler ile e-mailimizin body'sini oluşturuyoruz.
       */
      let body = `Atayan: ${atayan}\nİş: ${description}\nDeadline: ${deadline}`;
  

      /**
       * Farklı iş atamalarındaki e-mailler aynı thread altında gelmesin diye
       * her e-mail için farklı bir kod oluşturuyoruz.
       */
      let kod = Math.floor(Math.random() * 1000000)
  
      /**
       * Üstte cell'in checkbox olmasını ve değerinin true olmasını kontrol etmiştik.
       * Eğer bu koşullar sağlanıyorsa ve cell N sütununda ise if bloğuna giriyoruz.
       */
      if(cell == 'N') {

        /**
         * Üstte oluşturduğumuz kod ile birlikte e-mailimizin konusunu oluşturuyoruz.
         * Örnek: Tarafınıza yeni iş atandı! 123456
         */
        let subject = `Tarafınıza yeni iş atandı! ${kod}`;

        /**
         * Eğer adres Gökay Bağrıyanık ise, e-maili president@esnturkey.org'a gönderiyoruz.
         */
        if(address == "Gökay Bağrıyanık") {
          let email = "president@esnturkey.org";
          /**
           * E-mail içeriğine gerekli bilgileri ekliyoruz.
           */
          MailApp.sendEmail(email, subject, body, {
            cc: getCc(assigner)
          });
        }

        /**
         * Eğer adres Merve Ceylan ise, e-maili projectmanager@esnturkey.org'a gönderiyoruz. 
        */
        if(address == "Merve Ceylan") {
          let email = "projectmanager@esnturkey.org";
          MailApp.sendEmail(email, subject, body, {
            cc: getCc(assigner)
          });
        }
        
  
        /**
         * Eğer adres Gözde Özel ise, e-maili nr@esnturkey.org'a gönderiyoruz.
         */
        if(address == "Gözde Özel") {
          let email = "nr@esnturkey.org";
          MailApp.sendEmail(email, subject, body, {
            cc: getCc(assigner)
          });
        }
        
        /**
         * Eğer adres Nisa Gökyıldız ise, e-maili communication@esnturkey.org'a gönderiyoruz.
         */
        if(address == "Nisa Gökyıldız") {
          let email = "communication@esnturkey.org";
          MailApp.sendEmail(email, subject, body, {
            cc: getCc(assigner)
          });
        }

        /**
         * Eğer adres Doğukan Berk Demirdelen ise, e-maili treasurer@esnturkey.org'a gönderiyoruz. 
         */
        if(address == "Doğukan Berk Demirdelen") {
          let email = "treasurer@esnturkey.org";
          MailApp.sendEmail(email, subject, body, {
            cc: getCc(assigner)
          });
        }
  
        /**
         * Eğer adres Kaan Can Yıldırım ise, e-maili vicepresident@esnturkey.org'a gönderiyoruz.
         */
        if(address == "Kaan Can Yıldırım") {
          let email = "vicepresident@esnturkey.org";
          MailApp.sendEmail(email, subject, body, {
            cc: getCc(assigner)
          });
        }
  
        /**
         * Eğer adres Furkan Uçar ise, e-maili wpa@esnturkey.org'a gönderiyoruz.
         */
        if(address == "Furkan Uçar") {
          let email = "wpa@esnturkey.org";
          MailApp.sendEmail(email, subject, body, {
            cc: getCc(assigner)
          });
        }
  
        /**
         * Eğer adres Board ise, e-maili board@esnturkey.org'a gönderiyoruz.
         */
        if(address == "Board") {
          let email = "board@esnturkey.org";
          MailApp.sendEmail(email, subject, body, {
            cc: getCc(assigner)
          });
        }
  
        /**
         * Eğer hiçbiri değilse, hata mesajı veriyoruz.
         */
        else {
        Logger.log('The selected cell is not a checkbox or it is unchecked');
        }
      }
      
      /**
       * Eğer cell O sütununda ise, Denetleme Kuruluna e-mail gönderiyoruz. Prosedür aynı.
       */
      if (cell == 'O') {
        let email = 'auditors@esnturkey.org';
        let subject = `Yönetim Kurulunda bir iş ataması gerçekleştirildi! ${kod}`;
        let body = `Sevgili Denetleme Kurulu, \nDashboardda bir iş ataması gerçekleştirilmiştir. Bilginize sunarım. \nİş: ${description}\nAtayan: ${atayan} \nAtanan: ${address}\nDeadline:${deadline}\nİyi çalışmalar dilerim.\n\nNB Mail Bottan sevgilerle,`
        MailApp.sendEmail(email, subject, body, {
          cc: getCc(assigner)
        });
  
      }
    }

    /**
     * Burada ise, eğer cell M sütununda ise ve değeri 'Closed' ise if bloğuna giriyoruz.
     * Bu if bloğu, işin kapatılması durumunda çalışıyor.
     * İşin kapatılması durumunda, işin bilgilerini yeni bir satıra (aşağıdaki Closed kısmına) kopyalıyoruz.
     * Daha sonra, eski satırı siliyoruz.
     * Bu sayede kapatılan bir iş aktif işler arasında kalmıyor, Closed kısmına taşınıyor.
     */
    if (cell == 'M' && status == 'Closed') {
        let row = 170;
        let columnC = sheet.getRange("C" + row).getValue();
    
        while (columnC !== '') {
        row++;
        columnC = sheet.getRange("C" + row).getValue();
        }
    
        let addressCell = sheet.getRange('J' + number);
        let descriptionCell = sheet.getRange('C' + number);
        let assignerCell = sheet.getRange('B' + number);
        let deadlineCell = sheet.getRange('L' + number);
        let startdateCell = sheet.getRange('K' + number);
        let statusCell = sheet.getRange('M' + number);
    
        let addressValidation = addressCell.getDataValidation();
        let descriptionValidation = descriptionCell.getDataValidation();
        let assignerValidation = assignerCell.getDataValidation();
        let deadlineValidation = deadlineCell.getDataValidation();
        let startdateValidation = startdateCell.getDataValidation();
        let statusValidation = statusCell.getDataValidation();
    
        let targetAddressCell = sheet.getRange('J' + row);
        let targetDescriptionCell = sheet.getRange('C' + row);
        let targetAssignerCell = sheet.getRange('B' + row);
        let targetDeadlineCell = sheet.getRange('L' + row);
        let targetStartdateCell = sheet.getRange('K' + row);
        let targetStatusCell = sheet.getRange('M' + row);
    
        let value = descriptionCell.getValue();
        targetDescriptionCell.setValue(value);
        sheet.getRange('C' + row + ':I' + row).merge();
    
        let addressValue = addressCell.getValue();
        targetAddressCell.setValue(addressValue);
    
        let assignerValue = assignerCell.getValue();
        targetAssignerCell.setValue(assignerValue);
    
        let deadlineValue = deadlineCell.getValue();
        targetDeadlineCell.setValue(deadlineValue);
    
        let startdateValue = startdateCell.getValue();
        targetStartdateCell.setValue(startdateValue);
    
        let statusValue = statusCell.getValue();
        targetStatusCell.setValue(statusValue);
    
        targetAddressCell.setDataValidation(addressValidation);
        targetAssignerCell.setDataValidation(assignerValidation);
        targetDeadlineCell.setDataValidation(deadlineValidation);
        targetStartdateCell.setDataValidation(startdateValidation);
        targetStatusCell.setDataValidation(statusValidation);
        targetDescriptionCell.setDataValidation(descriptionValidation);
    
        sheet.deleteRow(number);
    }
  }
  
  
  /**
   * E-maili yollayacağımız kişiyle aynı mantıkta çalışıyor.
   * CC'yle kime e-mail yollayacağımızı belirliyoruz.
   */
  function getCc(assigner) {
    let email;
  
    if (assigner == "Gökay Bağrıyanık") {
      email = "president@esnturkey.org";
    } else if (assigner == "Merve Ceylan") {
      email = "projectmanager@esnturkey.org";
    } else if (assigner == "Gözde Özel") {
      email = "nr@esnturkey.org";
    } else if (assigner == "Nisa Gökyıldız") {
      email = "communication@esnturkey.org";
    } else if (assigner == "Doğukan Berk Demirdelen") {
      email = "treasurer@esnturkey.org";
    } else if (assigner == "Kaan Can Yıldırım") {
      email = "vicepresident@esnturkey.org";
    } else if (assigner == "Furkan Uçar") {
      email = "wpa@esnturkey.org";
    } else {
      email = "";
    }
  
    return email;
  }
```
