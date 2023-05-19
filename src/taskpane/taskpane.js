/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Define a new AngularJS module named "MsSignIn" with a dependency on "ngMaterial"
var app = angular.module("MsSignIn", ["ngMaterial"], function ($mdThemingProvider) {
  $mdThemingProvider.theme("default").primaryPalette("blue", { default: "500" }); 
  // Select the default theme then Set the primary color palette to blue with a default shade of 500
});

//define the controller of the Angular.js app 'SignInCtrl', which is associated with' MsSignIn' which is defined above 
app.controller("SignInCtrl", function ($scope) {
  var token = "";
  Office.onReady(function () {
   
    window.localStorage.setItem("SMS", "false");

    //retrieve token from local storage
    var accessToken = window.localStorage.getItem("accessToken");

    //checking if token exists
    if (accessToken != null) {
      var tokenTime = window.localStorage.getItem("time");
      tokenTime = new Date(tokenTime);
      var currentTime = new Date();
      
      //check if token is expired case 1
      if (currentTime.getDate() > tokenTime.getDate()) {
        //if token is expired, show log in dialog
        $scope.Login_Btn = true;
        $scope.MainPage = false;

        //if the value of $$phase is null, the $scope.$apply() method is called to ensure that the data is always in sync
        if (!$scope.$$phase) {
          $scope.$apply();
        }
      }
      
      //if token is expired case 2(after 60 minutes)
      else {
        var minutesDiff = currentTime.getTime() - tokenTime.getTime();
        var minutes = Math.floor(minutesDiff / 60000);
        if (minutes > 60) {
        //if session is expired, show log in dialog
          $scope.Login_Btn = true;
          $scope.MainPage = false;
          
        //if the value of $$phase is null, the $scope.$apply() method is called to ensure that the data is always in sync
          if (!$scope.$$phase) {
            $scope.$apply();
          }
        }

        //if session is not expired, show main page 
        else {
          $scope.Login_Btn = false;
          $scope.MainPage = true;
        
        //if the value of $$phase is null, the $scope.$apply() method is called to ensure that the data is always in sync
          if (!$scope.$$phase) {
            $scope.$apply();
          }
        }
      }
    }
    
    //if token is null (it does not exist)
    else { 
      //display log in dialogg
      $scope.Login_Btn = true;

       //if the value of $$phase is null, the $scope.$apply() method is called to ensure that the data is always in sync
      if (!$scope.$$phase) {
        $scope.$apply();
      }
    }

    //Callback function for dialog to store token
    function LogprocessMessage(arg) {
      Logindialog.close();
      token = arg.message;
      window.localStorage.setItem("accessToken", token);
      $scope.Login_Btn = false;
      window.localStorage.setItem("time", new Date());

      if (!$scope.$$phase) {
        $scope.$apply();
      }
    }

    $scope.getEmail = function (e) {
      window.localStorage.setItem("ToEmail", $scope.EmailAddress);
    };

    $scope.getMobile = function () {
      window.localStorage.setItem("ToPhone", "+" + $scope.mobile);
      console.log(window.localStorage.getItem("ToPhone"));
    };

    $scope.getSideNote = function () {
      window.localStorage.setItem("SideNote", $scope.sideNote);

      console.log(window.localStorage.getItem("SideNote"));
    };

    $scope.smsSend = function () {
      window.localStorage.setItem("SMS", $scope.sms);
    };

    var Logindialog;
    //Signing in with Microsoft
    $scope.SignInMS = function () {
      ////////////////Correct////////////////
      var link =
        "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=YOURCLIENTID&response_type=token&redirect_uri=https://localhost:3000/commands.html&scope=Mail.Send&response_mode=fragment&state=12345&nonce=678910";
      //Displaying office dialouge for login
      Office.context.ui.displayDialogAsync(link, { height: 50, width: 30 }, function (asyncResult) {
        Logindialog = asyncResult.value;
        Logindialog.addEventHandler(Office.EventType.DialogMessageReceived, LogprocessMessage);
      });
    };

    $scope.action = function (event) {

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
                          //
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

                          //TWILIO SID and KEY  (ENTER YOUR OWN)
                          var SID = "AC63eccee33222251cdddsda62a6331sdasdab3a3eef6711745c"
                          var Key = "0baa77ec3ffdffsdf2c79evvvva87142231242324d6f166f2d72"


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

                                  },
                                  error: function (data) {
                                    //If phone number is not valid, it will show the message
                                    console.log(response);
                                  }
                                });
                              }
                              else {

                              }
                            })
                            .fail(function (error) {
                              //if email is failed, it will tell to get token again

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
            }
          }
        }
      }
      else {
      }
      // Be sure to indicate when the add-in command function is complete
    }
  });
});



