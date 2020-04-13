import RPA from 'ts-rpa';
import { Router } from 'express';
import { TargetLocator, WebElement } from 'selenium-webdriver';

const SheetID = process.env.ABEMA_Report_SheetID;
const SheetName = `1週前`;

// エラー文を格納
const ErrorText = [];
// タブローの復帰数のデータを格納
const tableauData = [];

async function Start() {
  try {
    // デバッグログを最小限にする
    RPA.Logger.level = 'INFO';
    // 実行前にダウンロードフォルダを全て削除する
    await RPA.File.rimraf({ dirPath: `${process.env.WORKSPACE_DIR}` });
    await RPA.WebBrowser.get(process.env.ABEMA_Report_tableauURL);
    await CassoLogIN_function();
    await RPA.sleep(3000);
    await tableauOperation_function();
    await ReadCSV_function();
    await SetValue_function();
  } catch (Error) {
    ErrorText[0] = Error;
  } finally {
    RPA.Logger.info(ErrorText);
    await RPA.WebBrowser.quit();
  }
}

Start();

async function CassoLogIN_function() {
  try {
    const IDinput = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        id: 'username',
      }),
      5000
    );
    await RPA.WebBrowser.sendKeys(IDinput, [process.env.CyNumber]);
    const PWInput = await RPA.WebBrowser.findElementById(`password`);
    await RPA.WebBrowser.sendKeys(PWInput, [process.env.CyPass]);
    const Button = await RPA.WebBrowser.findElementByCSSSelector(
      `body > div > div.ping-body-container > div > form > div.ping-buttons > a`
    );
    await RPA.WebBrowser.mouseClick(Button);
  } catch (Error) {
    ErrorText[0] = Error;
  }
}

// フレームスイッチ
async function tableauOperation_function() {
  try {
    const Ifream = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        css: '#viz > iframe',
      }),
      10000
    );
    await RPA.WebBrowser.switchToFrame(Ifream);
    const StopButton = await RPA.WebBrowser.findElementByClassName(
      'tabToolbarButton tab-widget updates'
    );
    // 速度上げるために、タブローの更新ボタンを止める
    await StopButton.click();
    await RPA.Logger.info('【タブロー】更新止めました');
    await RPA.sleep(3000);
    // ステータス の (すべて) にチェックの外す
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByClassName('CFInnerContainer tab-ctrl-formatted-text tiledContent')[0].children[0].children[1].children[0].click();`
    );
    await RPA.sleep(2000);
    // ステータスの Abema復帰 のチェックを入れる
    const StatusList: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('CategoricalFilterBox')[0].children[2].children[1].children[1].children[0].children`
    );
    for (let i in StatusList) {
      const StatusText = await StatusList[i].getText();
      if (StatusText == `Abema復帰`) {
        console.log(StatusText);
        await RPA.WebBrowser.driver.executeScript(
          `document.getElementsByClassName('CategoricalFilterBox')[0].children[2].children[1].children[1].children[0].children[${i}].children[1].children[0].click()`
        );
        await RPA.sleep(1000);
        await RPA.Logger.info(
          `【タブロー】ステータス ${StatusText} にチェックしました`
        );
        break;
      }
    }
    // デバイス の全てにチェックを入れる
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByClassName('CFInnerContainer tab-ctrl-formatted-text tiledContent')[1].children[0].children[1].children[0].click()`
    );
    await RPA.sleep(1000);
    // デバイス の全てにチェックを外す
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByClassName('CFInnerContainer tab-ctrl-formatted-text tiledContent')[1].children[0].children[1].children[0].click()`
    );
    await RPA.sleep(1000);
    // デバイス SPTab-App のチェックを入れる
    const DeviceList: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('CFInnerContainer tab-ctrl-formatted-text tiledContent')[1].children[1].children[0].children`
    );
    for (let i in DeviceList) {
      const DeviceText = await DeviceList[i].getText();
      if (DeviceText == `SPTab-App`) {
        console.log(DeviceText);
        await RPA.WebBrowser.driver.executeScript(
          `document.getElementsByClassName('CFInnerContainer tab-ctrl-formatted-text tiledContent')[1].children[1].children[0].children[${i}].children[1].children[0].click()`
        );
        await RPA.Logger.info(
          `【タブロー】デバイス ${DeviceText} チェックしました`
        );
        await RPA.sleep(1000);
        break;
      }
    }
    // 画面右側の”メジャーネーム”の全てに一度チェック入れてそのあと外す
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByClassName('CFInnerContainer tab-ctrl-formatted-text tiledContent')[2].children[0].children[1].children[0].click()`
    );
    await RPA.sleep(1000);
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByClassName('CFInnerContainer tab-ctrl-formatted-text tiledContent')[2].children[0].children[1].children[0].click()`
    );
    await RPA.sleep(1000);
    const MajorNameList: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('CFInnerContainer tab-ctrl-formatted-text tiledContent')[2].children[1].children[0].children`
    );
    for (let i in MajorNameList) {
      const MajorNameText = await MajorNameList[i].getText();
      if (MajorNameText == `09.訪問UU`) {
        console.log(MajorNameText);
        await RPA.WebBrowser.driver.executeScript(
          `document.getElementsByClassName('CFInnerContainer tab-ctrl-formatted-text tiledContent')[2].children[1].children[0].children[${i}].children[1].children[0].click()`
        );
        await RPA.Logger.info(
          `【タブロー】メジャーネーム ${MajorNameText} にチェック入れました`
        );
        await RPA.sleep(1000);
        break;
      }
    }
    // 日付を変更する
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByClassName('RelativeDateFilter')[0].children[0].children[1].children[0].click()`
    );
    await RPA.sleep(2000);
    // 日付ダイアログの週を押す
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByName('week')[0].click()`
    );
    await RPA.sleep(1000);
    // 前週 にチェックを入れる
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByClassName('RelativeDateFilterDialog')[0].children[2].children[0].children[0].click()`
    );
    // 今週 にチェックを入れる
    /*
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByClassName('rradio')[2].click()`
    );
    */
    await RPA.sleep(2000);
    // 更新ボタンをおす
    const RefreshButton = await RPA.WebBrowser.findElementByClassName(
      'tabToolbarButton tab-widget refresh'
    );
    await RefreshButton.click();
    await RPA.sleep(5000);
    // 一度タブローのキャンバスをクリックする
    const CanvasElement: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('tab-clip')[0].children[1].children[0]`
    );
    await CanvasElement.click();
    await RPA.sleep(1000);
    // データダウンロード ボタンクリック
    const DLButton: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('tab-nonVizItems tab-fill-right showLabels')[0].children[4].children[1]`
    );
    await DLButton.click();
    await RPA.sleep(2000);
    // データダウンロード 実行
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByClassName('tab-downloadDialog')[0].children[3].click()`
    );
    await RPA.Logger.info('【タブロー】CSVダウンロード中...');
    await RPA.sleep(10000);
  } catch (Error) {
    ErrorText[0] = Error;
  }
}

// CSVを読み込む
async function ReadCSV_function() {
  try {
    var FileNames: any = await RPA.File.list();
    for (let i in FileNames) {
      if (FileNames[i].includes('.csv') == true) {
        FileNames = FileNames[i];
        break;
      }
    }
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
    console.log(SheetData[4]);
    const NewData = [];
    for (let i in SheetData[4]) {
      if (SheetData[4][i] == '09.訪問UU') {
        continue;
      }
      if (SheetData[4][i] == '全体') {
        continue;
      }
      if (SheetData[4][i] == 'SPTab-App') {
        continue;
      }
      if (SheetData[4][i] == 'Abema復帰') {
        continue;
      }
      tableauData.push(SheetData[4][i]);
    }
  } catch (Error) {
    ErrorText[0] = Error;
  }
}

// スプレッドシートに貼り付ける関数
async function SetValue_function() {
  await RPA.Logger.info('＊＊＊スプレッドシート貼り付けに移行＊＊＊');
  console.log(tableauData);
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });

  for (let i in tableauData) {
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: SheetID,
      range: `${SheetName}!K${Number(i) + 2}:K${Number(i) + 2}`,
      values: [[tableauData[i]]],
      parseValues: true,
    });
  }
}
