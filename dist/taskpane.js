/******/ (() => { // webpackBootstrap
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    document.addEventListener('DOMContentLoaded', function () {
      loadSettings();
      document.getElementById('saveBtn').addEventListener('click', saveSettings);
      document.getElementById('testBtn').addEventListener('click', testCheck);
    });
  }
});
function loadSettings() {
  try {
    var savedSettings = localStorage.getItem('emailCheckSettings');
    if (savedSettings) {
      var settings = JSON.parse(savedSettings);
      document.getElementById('subjectCheck').checked = settings.subjectCheck || false;
      document.getElementById('recipientCheck').checked = settings.recipientCheck || false;
      document.getElementById('attachmentCheck').checked = settings.attachmentCheck || false;
      document.getElementById('keywordCheck').checked = settings.keywordCheck || false;
      document.getElementById('externalDomainCheck').checked = settings.externalDomainCheck || false;
      document.getElementById('attachmentKeywords').value = settings.attachmentKeywords || '添付,attachment,ファイル';
      document.getElementById('warningKeywords').value = settings.warningKeywords || '機密,重要,urgent,極秘';
      document.getElementById('internalDomains').value = settings.internalDomains || '';
    }
  } catch (error) {
    showResult('設定の読み込みに失敗しました: ' + error.message, 'error');
  }
}
function saveSettings() {
  try {
    var settings = {
      subjectCheck: document.getElementById('subjectCheck').checked,
      recipientCheck: document.getElementById('recipientCheck').checked,
      attachmentCheck: document.getElementById('attachmentCheck').checked,
      keywordCheck: document.getElementById('keywordCheck').checked,
      externalDomainCheck: document.getElementById('externalDomainCheck').checked,
      attachmentKeywords: document.getElementById('attachmentKeywords').value,
      warningKeywords: document.getElementById('warningKeywords').value,
      internalDomains: document.getElementById('internalDomains').value
    };
    localStorage.setItem('emailCheckSettings', JSON.stringify(settings));
    showResult('設定を保存しました', 'success');
  } catch (error) {
    showResult('設定の保存に失敗しました: ' + error.message, 'error');
  }
}
function testCheck() {
  try {
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
                    var issues = performChecks(subject, recipients, body, attachments);
                    displayTestResults(issues);
                  } else {
                    showResult('添付ファイル情報の取得に失敗しました', 'error');
                  }
                });
              } else {
                showResult('メール本文の取得に失敗しました', 'error');
              }
            });
          } else {
            showResult('宛先情報の取得に失敗しました', 'error');
          }
        });
      } else {
        showResult('件名の取得に失敗しました', 'error');
      }
    });
  } catch (error) {
    showResult('テスト実行中にエラーが発生しました: ' + error.message, 'error');
  }
}
function performChecks(subject, recipients, body, attachments) {
  var settings = getSettings();
  var issues = [];
  if (settings.subjectCheck && (!subject || subject.trim() === '')) {
    issues.push('件名が空です');
  }
  if (settings.recipientCheck && recipients.length === 0) {
    issues.push('宛先が設定されていません');
  }
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
function getSettings() {
  try {
    var savedSettings = localStorage.getItem('emailCheckSettings');
    if (savedSettings) {
      return JSON.parse(savedSettings);
    }
  } catch (error) {
    console.error('設定の読み込みエラー:', error);
  }
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
function displayTestResults(issues) {
  if (issues.length === 0) {
    showResult('✓ チェック完了: 問題は見つかりませんでした', 'success');
  } else {
    var message = '⚠ 以下の問題が見つかりました:\n' + issues.map(function (issue) {
      return '• ' + issue;
    }).join('\n');
    showResult(message, 'error');
  }
}
function showResult(message, type) {
  var resultDiv = document.getElementById('result');
  resultDiv.textContent = message;
  resultDiv.className = 'result ' + type;
  setTimeout(function () {
    resultDiv.textContent = '';
    resultDiv.className = '';
  }, 5000);
}
/******/ })()
;
//# sourceMappingURL=taskpane.js.map