const setMenu = ({
  name = "メニュー",
  items
}) => {
  const ui  = SpreadsheetApp.getUi();
  const menu = ui.createMenu(name);

  for (const menuItem of items) {
    menu.addItem(menuItem.name, menuItem.funcName);
  }

  menu.addToUi();
}

const outErrorMessage = (message) => {
  console.log(message);
  Browser.msgBox(message);
};

const replaceBody = (row, body) => {
  return body.replace("{company}", row[0]);
};

class MailSenderSheet {
  constructor(ss) {
    this.sheet = ss.getSheetByName("メールテンプレート");
    this.from = null;
    this.alias = null;
    this.subject = null;
    this.body = null;
    if (this.sheet !== null) {
      this.from = this.sheet.getRange("B1").getValue();
      this.alias = this.sheet.getRange("B2").getValue();
      this.subject = this.sheet.getRange("B3").getValue();
      this.body = this.sheet.getRange("B4").getValue();
    }
  }

  getOption() {
    const mailOption = {
      name: this.alias,
      from: this.from
    };

    if (mailOption.name === "") {
      delete mailOption.name;
    }

    if (mailOption.from === "") {
      delete this.from;
    }

    return mailOption;
  }
}

class ToListSheet {
  constructor(ss) {
    this.sheet = ss.getSheetByName("メール宛先リスト");
    this.list = new Array;
    if (this.sheet !== null) {
      this.list = this.sheet.getRange(2, 1, this.sheet.getLastRow(), this.sheet.getLastColumn())
        .getValues()
        .filter(row => row.indexOf("") === -1);
    }
  }
}

function onOpen()  {
  setMenu({
    items: [{
      name: "問合せ先取得",
      funcName: "getInquiryList"
    }, {
      name: "メール送信",
      funcName: "sendInquiryMail"
    }]
  });
}

function sendInquiryMail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sender = new MailSenderSheet(ss);
  const list = new ToListSheet(ss).list;

  if (sender.from === "" || sender.subject === "" ||sender.body === "") {
    return outErrorMessage("メールを作成できません。メール本文シートの各項目に入力してください。");
  }

  if (list.length === 0 || list.length > 30) {
    return outErrorMessage("メールの宛先リストに有効なデータを1～30件設定してください。");
  }

  list.forEach(function(row) {
    GmailApp.sendEmail(row[1], sender.subject, replaceBody(row, sender.body), sender.getOption());
  });
}

function getInquiryList() {

}
