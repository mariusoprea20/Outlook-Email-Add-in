Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {


  //Send Email
  var token = window.localStorage.getItem("accessToken");

  if (token != null) {

    var tokenTime = window.localStorage.getItem("time");
    tokenTime = new Date(tokenTime);
    var currentTime = new Date();


    // Checking if token is expired or not

    if (currentTime.getDate() > tokenTime.getDate()) {
      //  if token is expired , it will show notification and send button will be disbaled
      mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Session is expired, Please Sign in again in the add-in' });
      //Disabling on send button
      event.completed({ allowEvent: false });


    }

    else {
      var minutesDiff = currentTime.getTime() - tokenTime.getTime();
      var minutes = Math.floor(minutesDiff / 60000);
      if (minutes > 60) {

        //  if token is expired , it will show notification and send button will be disbaled
        mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Session is expired, Please Sign in again in the add-in' });
        //Disabling on send button
        event.completed({ allowEvent: false });

      }

      else {

        if (window.localStorage.getItem("ToEmail") != undefined) {

          //If token is not expired


          var token = window.localStorage.getItem("accessToken");
          var Message = "";
          var subject = "";

          //Getting from addresss
          Office.context.mailbox.item.from.getAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
              Message = "From: " + asyncResult.value.emailAddress;

              //Getting subject of email
              Office.context.mailbox.item.subject.getAsync(function (Result) {
                if (Result.status == Office.AsyncResultStatus.Succeeded) {
                  Message = " Subject: " + Result.value + " " + Message;

                  // getting to email address
                  Office.context.mailbox.item.to.getAsync(function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                      var ToEmails = " To :";
                      //if multiple email addresses in To:
                      for (let i = 0; i < asyncResult.value.length; i++) {
                        ToEmails += asyncResult.value[i].emailAddress + ",";
                      }
                      Message = Message + ToEmails;
                      var toEmail = window.localStorage.getItem("ToEmail");
                      var sideNote = window.localStorage.getItem("SideNote");

                      //Body for email sending
                      var PostBody = {

                        "message": {
                          subject: "Subject",
                          body: {
                            contentType: "text",
                            content: Message + " Side Note:" + sideNote,
                          },
                          toRecipients: [
                            {
                              emailAddress: {
                                address: toEmail,
                              },
                            },
                          ],
                        },
                      };


                      //TWILIO SID and KEY (ENTER YOUR OWN
                      var SID = "ER53eccee3dsdasdas51cd62a633gfgedfg894384fudfs1b3a3eef67115c"
                      var Key = "7ec3dfsdfsdjfsdsdnfsdfkl2324235232f2c79e87142d6f1009u432fdhsfh"



                      //Email sending ajax call
                      var settings = {
                        url: "https://graph.microsoft.com/v1.0/me/sendMail",
                        method: "POST",
                        timeout: 0,
                        headers: {
                          "Content-Type": "application/json",
                          "Authorization": "bearer " + token,
                        },
                        data: JSON.stringify(PostBody),
                      };

                      $.ajax(settings)
                        .done(function (response) {


                          if (!(window.localStorage.getItem("SMS") == "true")) {
                            //If everything is done, email will be sent after clicking on-send button
                            $.ajax({
                              type: 'POST',
                              url: 'https://api.twilio.com/2010-04-01/Accounts/' + SID + '/Messages.json',
                              data: {
                                "To": window.localStorage.getItem("ToPhone"),
                                "From": "+447883317644",
                                "Body": Message + " Side Note:" + sideNote,
                              },
                              beforeSend: function (xhr) {
                                xhr.setRequestHeader("Authorization", "Basic " + btoa(SID + ':' + Key));
                              },
                              success: function (data) {
                                event.completed({ allowEvent: true });
                              },
                              error: function (data) {
                                //If phone number is not valid, it will show the message
                                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Phone number is not valid' });
                                //Disabling on-send button
                                event.completed({ allowEvent: false });
                              }
                            });
                          }

                          else {
                            event.completed({ allowEvent: true });
                          }
                        })
                        .fail(function (error) {
                          //if email is failed, it will tell to get token again
                          mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Session is expired, Please Sign in again in the add-in' });
                          //Disabling on-send button
                          event.completed({ allowEvent: false });
                        });
                    }
                  });
                }
              });
            }
          });
        }
        else {
          //  if token is expired , it will show notification and send button will be disbaled
          mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Session is expired, Please Sign in again in the add-in' });
          //Disabling on send button
          event.completed({ allowEvent: false });
        }
      }
    }
  }
  else {
    //  if token is expired , it will show notification and send button will be disbaled
    mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Session is expired, Please Sign in again in the add-in' });
    //Disabling on send button
    event.completed({ allowEvent: false });
  }

  // Be sure to indicate when the add-in command function is complete
}
function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}
const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;

