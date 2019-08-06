// Auto-Confirmation Email for courses

function AutoConfirmation(e) {

    try {

      var theirEmail, subject, message;
      var body, listing, signature;
      var ourName, firstName, surName;

      //Name variables
      ourName = "Helpdesk Unil";
      theirEmail = e.namedValues["Email @unil.ch"].toString();
      firstName = e.namedValues["Prénom"].toString();
      surName = e.namedValues["Nom de famille"].toString();

      //Object
      subject = "Confirmation d'inscription - cours info étus";

      //Signature
      signature =
          "<br>Yvan Uhlmann" +
          "<br>||||||||||||||||||||||||||||||||||||||||||||||" +
          "<br>Centre informatique UNIL" +
          "<br>0041 21 692 2211 |  https://unil.ch/ci/cours-etudiants";

      //Body
      //Test if the Email finishes with '@unil.ch'
      if ( theirEmail.indexOf('@unil.ch') !== -1 ) {
        body =
          firstName + " " + surName + "," +
          "<br><br>Je confirme votre inscription aux cours informatiques pour étudiants." +
            "<br>Dû à la quantité forte de demandes d'inscriptions, merci d'anoncer toute absence au plus vite."
      }
      else {
        body = "Vous ne vous êtes pas enregisté avec une adresse mail de l'Unil." +
          "<br>Votre inscription n'a donc pas été prise en compte."
      }

      //List the user inscriptions
      var ss, columns;
      ss = SpreadsheetApp.getActiveSheet();
      columns = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];

      listing = "<br>------------<br>Vos inscriptions:<br>";
      for ( var keys in columns ) {
          var key = columns[keys];
          if ( e.namedValues[key] && e.namedValues[key] != "" && (key == "Cours Unil" || key == "Cours Systèmes d'exploitation" || key == "Cours bureautique" || key == "Cours spécialistes") )
          {listing += key + ' : '+ e.namedValues[key] + "<br>";
          }
      }

      //Create the whole mail
      message = body + "<br>" + signature + "<br><br>" + listing;

      //Send Email
      var cosmetics = {name: ourName, htmlBody: message};
      GmailApp.sendEmail(theirEmail, subject, message, cosmetics );

    } catch (e) {
        Logger.log(e.toString());
    }

}
