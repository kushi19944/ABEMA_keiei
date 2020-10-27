import RPA from 'ts-rpa';
import { TargetLocator, WebElement } from 'selenium-webdriver';

const SheetID = process.env.ABEMA_Report_SheetID;
const SheetName = `全データ`;

// エラー文を格納
let ErrorText;
// タブローの復帰数のデータを格納
var tableauData = '';

async function Start() {
  try {
    // 実行前にダウンロードフォルダを全て削除する
    await RPA.File.rimraf({ dirPath: `${process.env.WORKSPACE_DIR}` });
    // タブロー復帰(DAU)
    await RPA.WebBrowser.get(process.env.ABEMA_Report_tableauURL_DAU);
    await CassoLogIN_function();
    await RPA.sleep(3000);
    await tableauOperation_function();
    await ReadCSV_function();
    await SetValue_function();

    // タブロー LAP
    // 実行前にダウンロードフォルダを全て削除する
    await RPA.File.rimraf({ dirPath: `${process.env.WORKSPACE_DIR}` });
    await RPA.WebBrowser.get(process.env.ABEMA_Report_tableauURL_LAP);
    try {
      await CassoLogIN_function();
    } catch {}
    await RPA.sleep(3000);
    await tableauOperation_function();
    await ReadCSV_function();
    await SetValue_LAP_function();
    await RPA.Logger.info(`取得データ:` + tableauData);
  } catch (Error) {
    ErrorText = Error;
    await RPA.WebBrowser.takeScreenshot();
  } finally {
    RPA.Logger.info(ErrorText);
    await RPA.WebBrowser.quit();
  }
}

Start();

async function CassoLogIN_function() {
  const IDinput = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      id: 'username',
    }),
    20000
  );
  await RPA.WebBrowser.sendKeys(IDinput, [process.env.CyNumber]);
  const PWInput = await RPA.WebBrowser.findElementById(`password`);
  await RPA.WebBrowser.sendKeys(PWInput, [process.env.CyPass]);
  const Button = await RPA.WebBrowser.findElementByCSSSelector(
    `body > div > div.ping-body-container > div > form > div.ping-buttons > a`
  );
  await RPA.WebBrowser.mouseClick(Button);
}

// フレームスイッチ
async function tableauOperation_function() {
  const Ifream = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      css: '#viz > iframe',
    }),
    10000
  );
  await RPA.WebBrowser.switchToFrame(Ifream);

  const DownloadButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      id: 'download-ToolbarButton',
    }),
    10000
  );
  // 一度タブローのキャンバスをクリックする
  const CanvasElement: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('tab-clip')[0].children[1].children[0]`
  );
  // タブローが更新されていない場合は終了させる
  if ((await CanvasElement.isDisplayed()) == false) {
    await RPA.Logger.info(`【タブロー】数値更新されていません.RPA停止します`);
    await RPA.WebBrowser.quit();
    await RPA.sleep(1000);
    process.exit(0);
  }

  // ダウンロードバー(上部)クリック
  await DownloadButton.click();
  await RPA.sleep(3000);
  RPA.Logger.info('ダウンロードバークリック');
  // 集計クリック
  await RPA.WebBrowser.driver.executeScript(
    `document.getElementById("DownloadDialog-Dialog-Body-Id").children[0].children[3].click()`
  );

  // データダウンロード 実行
  const DLbutton3 = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'fdiufnn low-density',
    }),
    10000
  );
  await DLbutton3.click();
  await RPA.Logger.info('【タブロー】CSVダウンロード中...');
  await RPA.sleep(10000);
}

// CSVを読み込む
async function ReadCSV_function() {
  const filelists = await RPA.File.list();
  const FileNames = [];
  console.log(filelists);
  for (let i in filelists) {
    if (filelists[i].includes('.csv') == true) {
      FileNames[0] = filelists[i];
      console.log(FileNames);
      break;
    }
  }
  if (FileNames.length == 1) {
    const SheetData = await RPA.CSV.read({
      filename: `${FileNames}`,
      encoding: 'UTF16',
      relaxColumnCount: true,
      delimiter: `\t`, //切り分ける文字を指定
    });
    /* エンコードの種類
        [
          'UTF32',
          'UTF16',
          'UTF16BE',
          'UTF16LE',
          'BINARY',
          'ASCII',
          'JIS',
          'UTF8',
          'EUCJP',
          'SJIS',
          'UNICODE',
        ];
        */
    console.log(SheetData);
    for (let i in SheetData[4]) {
      try {
        if (
          SheetData[4][i] == '09.訪問UU' ||
          SheetData[4][i] == '全体' ||
          SheetData[4][i] == 'SPTab-App' ||
          SheetData[4][i] == 'Abema復帰'
        ) {
          continue;
        }
        tableauData = SheetData[4][i];
      } catch {}
    }
  }
}

// スプレッドシートに貼り付ける関数
async function SetValue_function() {
  await RPA.Logger.info('＊＊＊スプレッドシート貼り付けに移行＊＊＊');
  console.log(tableauData);
  if (tableauData == '') {
    RPA.Logger.info('タブローのデータがありません。停止します');
    await RPA.WebBrowser.quit();
    await RPA.sleep(2000);
    await process.exit();
  }
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });
  const DayData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: SheetID,
    range: `${SheetName}!B1:B1`,
  });
  const FirstData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: SheetID,
    range: `${SheetName}!B2:B100`,
  });
  for (let i in FirstData) {
    if (FirstData[i][0] == DayData[0][0]) {
      console.log(Number(i) + 2);
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: SheetID,
        range: `${SheetName}!N${Number(i) + 2}:N${Number(i) + 2}`,
        values: [[tableauData]],
        parseValues: true,
      });
      break;
    }
  }
}

// スプレッドシートに貼り付ける関数
async function SetValue_LAP_function() {
  await RPA.Logger.info('＊＊＊スプレッドシート貼り付けに移行＊＊＊');
  console.log(tableauData);
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });
  const DayData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: SheetID,
    range: `${SheetName}!B1:B1`,
  });
  const FirstData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: SheetID,
    range: `${SheetName}!B2:B100`,
  });
  for (let i in FirstData) {
    if (FirstData[i][0] == DayData[0][0]) {
      console.log(Number(i) + 2);
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: SheetID,
        range: `${SheetName}!Q${Number(i) + 2}:Q${Number(i) + 2}`,
        values: [[tableauData]],
        parseValues: true,
      });
      break;
    }
  }
}
