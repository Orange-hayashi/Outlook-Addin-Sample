/******/ (() => { // webpackBootstrap
/*!**********************************!*\
  !*** ./src/commands/commands.js ***!
  \**********************************/
Office.onReady(function () {
  // register a function named checkBeforeSend
  Office.actions.associate("checkBeforeSend", checkBeforeSend);
});
function checkBeforeSend(event) {
  var settings = getSettings();

  // すべてのチェックが無効の場合は送信を許可
  if (!settings.subjectCheck && !settings.recipientCheck && !settings.attachmentCheck && !settings.keywordCheck && !settings.externalDomainCheck) {
    event.completed({
      allowEvent: true
    });
    return;
  }

  // メール情報を取得してチェック実行
  Office.context.mailbox.item.subject.getAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var subject = result.value || '';
      Office.context.mailbox.item.to.getAsync(function (toResult) {
        if (toResult.status === Office.AsyncResultStatus.Succeeded) {
          var recipients = toResult.value || [];
          Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
            if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
              var body = bodyResult.value || '';
              Office.context.mailbox.item.attachments.getAsync(function (attachResult) {
                if (attachResult.status === Office.AsyncResultStatus.Succeeded) {
                  var attachments = attachResult.value || [];
                  var issues = performChecks(subject, recipients, body, attachments, settings);
                  if (issues.length > 0) {
                    showConfirmationDialog(issues, event);
                  } else {
                    event.completed({
                      allowEvent: true
                    });
                  }
                } else {
                  event.completed({
                    allowEvent: true
                  });
                }
              });
            } else {
              event.completed({
                allowEvent: true
              });
            }
          });
        } else {
          event.completed({
            allowEvent: true
          });
        }
      });
    } else {
      event.completed({
        allowEvent: true
      });
    }
  });
}
function performChecks(subject, recipients, body, attachments, settings) {
  var issues = [];

  // 件名チェック
  if (settings.subjectCheck && (!subject || subject.trim() === '')) {
    issues.push('件名が空です');
  }

  // 宛先チェック
  if (settings.recipientCheck && recipients.length === 0) {
    issues.push('宛先が設定されていません');
  }

  // 添付ファイルチェック
  if (settings.attachmentCheck) {
    var keywords = settings.attachmentKeywords.split(',').map(function (k) {
      return k.trim();
    }).filter(function (k) {
      return k;
    });
    var hasAttachmentKeyword = keywords.some(function (keyword) {
      return subject.includes(keyword) || body.includes(keyword);
    });
    if (hasAttachmentKeyword && attachments.length === 0) {
      issues.push('添付ファイルに関するキーワードがありますが、ファイルが添付されていません');
    }
  }

  // キーワードチェック
  if (settings.keywordCheck) {
    var _keywords = settings.warningKeywords.split(',').map(function (k) {
      return k.trim();
    }).filter(function (k) {
      return k;
    });
    var foundKeywords = _keywords.filter(function (keyword) {
      return subject.includes(keyword) || body.includes(keyword);
    });
    if (foundKeywords.length > 0) {
      issues.push('注意すべきキーワードが含まれています: ' + foundKeywords.join(', '));
    }
  }

  // 外部ドメインチェック
  if (settings.externalDomainCheck && settings.internalDomains) {
    var internalDomains = settings.internalDomains.split(',').map(function (d) {
      return d.trim();
    }).filter(function (d) {
      return d;
    });
    var externalRecipients = recipients.filter(function (recipient) {
      var email = recipient.emailAddress;
      return !internalDomains.some(function (domain) {
        return email.includes(domain);
      });
    });
    if (externalRecipients.length > 0) {
      issues.push('外部ドメインへの送信が含まれています');
    }
  }
  return issues;
}
function showConfirmationDialog(issues, event) {
  var issueList = issues.map(function (issue) {
    return "<li>".concat(issue, "</li>");
  }).join('');
  var dialogHtml = "\n        <html>\n        <head>\n            <meta charset=\"UTF-8\">\n            <style>\n                body {\n                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;\n                    margin: 20px;\n                    background-color: #fff;\n                }\n                .header {\n                    color: #d13438;\n                    font-size: 18px;\n                    font-weight: bold;\n                    margin-bottom: 15px;\n                }\n                .issues {\n                    background-color: #fde7e9;\n                    border: 1px solid #d13438;\n                    border-radius: 4px;\n                    padding: 15px;\n                    margin-bottom: 20px;\n                }\n                ul {\n                    margin: 10px 0;\n                    padding-left: 20px;\n                }\n                li {\n                    margin-bottom: 5px;\n                }\n                .button-container {\n                    text-align: center;\n                    margin-top: 20px;\n                }\n                button {\n                    padding: 10px 20px;\n                    margin: 0 10px;\n                    border: none;\n                    border-radius: 4px;\n                    font-size: 14px;\n                    cursor: pointer;\n                }\n                .send-btn {\n                    background-color: #d13438;\n                    color: white;\n                }\n                .cancel-btn {\n                    background-color: #f3f2f1;\n                    color: #323130;\n                    border: 1px solid #d2d0ce;\n                }\n            </style>\n        </head>\n        <body>\n            <div class=\"header\">\u26A0 \u8AA4\u9001\u4FE1\u30C1\u30A7\u30C3\u30AF</div>\n            <div class=\"issues\">\n                <strong>\u4EE5\u4E0B\u306E\u554F\u984C\u304C\u898B\u3064\u304B\u308A\u307E\u3057\u305F:</strong>\n                <ul>".concat(issueList, "</ul>\n            </div>\n            <p>\u305D\u308C\u3067\u3082\u9001\u4FE1\u3057\u307E\u3059\u304B\uFF1F</p>\n            <div class=\"button-container\">\n                <button class=\"send-btn\" onclick=\"allowSend()\">\u9001\u4FE1\u3059\u308B</button>\n                <button class=\"cancel-btn\" onclick=\"cancelSend()\">\u30AD\u30E3\u30F3\u30BB\u30EB</button>\n            </div>\n            <script>\n                function allowSend() {\n                    Office.context.ui.messageParent(JSON.stringify({ allowEvent: true }));\n                }\n                function cancelSend() {\n                    Office.context.ui.messageParent(JSON.stringify({ allowEvent: false }));\n                }\n                Office.onReady(() => {\n                    // Dialog is ready\n                });\n            </script>\n        </body>\n        </html>\n    ");
  Office.context.ui.displayDialogAsync("data:text/html,".concat(encodeURIComponent(dialogHtml)), {
    height: 60,
    width: 50
  }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      var dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        try {
          var response = JSON.parse(arg.message);
          dialog.close();
          event.completed(response);
        } catch (error) {
          dialog.close();
          event.completed({
            allowEvent: false
          });
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
        dialog.close();
        event.completed({
          allowEvent: false
        });
      });
    } else {
      // ダイアログの表示に失敗した場合は送信を許可
      event.completed({
        allowEvent: true
      });
    }
  });
}
function getSettings() {
  try {
    var savedSettings = localStorage.getItem('emailCheckSettings');
    if (savedSettings) {
      return JSON.parse(savedSettings);
    }
  } catch (error) {
    console.error('設定の読み込みエラー:', error);
  }

  // デフォルト設定（すべて無効）
  return {
    subjectCheck: false,
    recipientCheck: false,
    attachmentCheck: false,
    keywordCheck: false,
    externalDomainCheck: false,
    attachmentKeywords: '添付,attachment,ファイル',
    warningKeywords: '機密,重要,urgent,極秘',
    internalDomains: ''
  };
}
/******/ })()
;
//# sourceMappingURL=commands.js.map