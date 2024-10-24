function mail_sender_invitation() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('guests'); // Access the 'guests' tab 
    if (!sheet) {
      throw new Error("The 'guests' sheet was not found.");
    }

    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5); // Get all relevant rows and columns
    var data = range.getValues();
    
    data.forEach(function(row, index) {
      try {
        var email = row[0];
        var subject = row[1];
        var name = row[2];

        if (!email || !subject || !name) {
          throw new Error(`Missing data in row ${index + 2}`);
        }

        var signature = `
          <br>
          <strong style="color:#191970;">"Your name""</strong><br><br>
          <span style="color:#191970;">For any questions contact me:<br>
          <span style="color:#191970;"><strong>Mobile:</strong></span> "Your phone number"<br>
          <span style="color:#191970;"><strong>E-mail:</strong></span> "Your email"<br>
        `;

        // Email body in HTML format
        var htmlBody = `
          <html lang="es">
          <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Celebration Invitation</title>
            <style>
              body {
                font-family: 'Arial', sans-serif;
                background-color: #f7f7f7;
                margin: 0;
                padding: 0;
              }

              .container {
                background-image: url('https://image.slidesdocs.com/responsive-images/docs/advertising-of-a-festive-blue-balloon-gift-for-a-birthday-party-design-page-border-background-word-template_fa0a0f7447__1131_1600.jpg'); /* Replace with your image link */
                background-size: cover;
                background-position: center;
                color: #fff;
                max-width: 700px;
                margin: 40px auto;
                padding-top: 125px;
                padding-bottom: 50px;
                padding-right: 50px;
                padding-left: 50px;
                border-radius: 10px;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
                text-align: center;
              }

              h1 {
                color: #191970;
                font-size: 28px;
                margin-bottom: 10px;
              }

              p {
                color: #000000;
                font-size: 16px;
                line-height: 1.6;
              }

              .event-details {
                background-color: rgba(240, 248, 255, 0.8);  /* Transparent */
                padding: 15px;
                margin: 20px 0;
                border-radius: 8px;
              }

              .event-details p {
                margin: 5px 0;
                font-size: 18px;
              }

              .cta-button {
                background-color: #191970;
                color: #F0FFFF;
                padding: 10px 20px;
                text-decoration: none;
                border-radius: 5px;
                font-size: 18px;
                display: inline-block;
                margin-top: 20px;
              }

              .cta-button:hover {
                background-color: #004c99;
                color: #F0F8FF;
              }

              .footer {
                margin-top: 30px;
                color: #000000;
                font-size: 16px;
              }

              .dress-code {
                background-color: rgba(240, 248, 255, 0.8);  /* Transparent */
                padding: 15px;
                margin: 20px 0;
                border-radius: 8px;
                text-align: left;
              }

              .dress-code h2 {
                color: #191970;
                font-size: 14px;
                margin-bottom: 10px;
              }

              .dress-code p {
                font-size: 13px;
                color: #333;
                margin: 5px 0;
              }

            </style>
          </head>
          <body>

            <div class="container">
              <h1>I would love to have you!</h1>

              <p>Dear <strong>${name}</strong>,</p>

              <p>As you know I'm turning 30 soon! And I would love to celebrate it with family and friends. I hope to have your presence:</p>

              <div class="event-details">
                <p><img src="https://cdn-icons-png.flaticon.com/512/876/876805.png" alt="Date" style="margin-bottom: -1px;" width="15"/>
                <strong>  Date:</strong> December 21, 2024</p>
                <p><img src="https://cdn-icons-png.flaticon.com/512/2089/2089758.png" alt="Time" style="margin-bottom: -1px;" width="15"/>
                <strong>  Time:</strong> from 5 pm </p>
                <p><img src="https://cdn-icons-png.flaticon.com/512/450/450016.png" alt="Location" style="margin-bottom: -1px;" width="15"/>
                <strong>  Location:</strong> "Add the address" <br>
              </div>
              <p>It will be an intimate and special gathering, and I hope you can join me on this day.</p>

              <p>Please confirm your attendance by <strong>"Inserte the date"</strong> by clicking on the following link and filling out the form:</p>

              <a href="Insert the link to your form" class="cta-button"><strong style="color:#F8F8FF;">Confirm attendance</strong></a>

              <p class="footer">If you have any questions or need more information, don't hesitate to contact me.<br>We look forward to seeing you soon!</p>

              <p class="footer">With love,<br>${signature}</p>

              <div class="dress-code">
                <h2>Important Notes:</h2>
                <p>- This is an adult-oriented celebration where alcohol will be served. I kindly ask that you don't bring children or minors.</p>
                <p>-<strong> Dress code:</strong> Elegant or smart casual.</p>
                <p>- Wear something comfortable but elegant to fully enjoy the celebration.</p>            
                <p>- Only those registered in the form will be admitted.</p>
              </div>
            </div>

          </body>
          </html>
        `;

        // Send the email with the attached file
        GmailApp.sendEmail(email, subject, '', {
          htmlBody: htmlBody
        });

        Logger.log(`Email sent to ${email}`);
      } catch (error) {
        Logger.log(`Error in row ${index + 2}: ${error.message}`);
      }
    });

  } catch (error) {
    Logger.log(`General error: ${error.message}`);
  }
}