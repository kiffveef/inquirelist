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

class SearchConfigSheet {
  constructor(ss) {
    this.sheet = ss.getSheetByName("config");
    this.apiKey = null;
    this.engineId = null;
    this.searchUrl = "https://www.googleapis.com/customsearch/v1?key=%s&cx=%s&q=%s&num=%s";
    if (this.sheet !== null) {
      this.apiKey = this.sheet.getRange("B1").getValue();
      this.engineId = this.sheet.getRange("B2").getValue();
    }
  }

  searchKeyword(keyword) {
    const url = Utilities.formatString(
      this.searchUrl,
      this.apiKey,
      this.engineId,
      encodeURI(Utilities.formatString("問い合わせ %s", keyword)),
      1
    );

    return JSON.parse(UrlFetchApp.fetch(url).getContentText("UTF-8")).items[0].link;
  }
}

class CompanySheet {
  constructor(ss) {
    this.sheet = ss.getSheetByName("企業リスト");
    this.list = new Array;
    if (this.sheet !== null) {
      this.list = this.sheet.getRange(1, 1, this.sheet.getLastRow(), this.sheet.getLastColumn())
        .getValues()
        .filter(row => (row[0] !== "" && row[1] === ""))
        .map(row => ({
          name: row[0],
          url: row[1],
          urlCell: this.sheet.createTextFinder(row[0]).findNext().offset(0, 1).getA1Notation()
        }));
    }
  }

  setUrl(a1cell, url) {
    this.sheet.getRange(a1cell).setValue(url);
    return;
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

  addMail(company, mails) {
    Array.from(new Set(mails)).forEach(mail => {
      this.sheet.appendRow([company, mail]);
    });
    return;
  }
}

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

function getInquiryList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customSearch = new SearchConfigSheet(ss);
  const companies = new CompanySheet(ss);
  const toList = new ToListSheet(ss);

  companies.list.forEach(company => {
    company.url = customSearch.searchKeyword(company.name);
    companies.setUrl(company.urlCell, company.url);

    var contents = UrlFetchApp.fetch(company.url).getContentText();
    var $ = Cheerio.load(contents, {
      decodeEntities: false
    });
    var matches = $("body")
      .first()
      .text()
      .match(/[a-zA-Z0-9]+([\._-]?[a-zA-Z0-9])*@[a-zA-Z0-9]+([\.-]?[a-zA-Z0-9])*(\.\w{2,3})/gi);
    if (matches !== null) {
      toList.addMail(company.name, matches);
    }
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

  list.forEach(row => {
    GmailApp.sendEmail(row[1], sender.subject, replaceBody(row, sender.body), sender.getOption());
  });
}