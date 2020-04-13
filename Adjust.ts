import RPA from 'ts-rpa';
import { TargetLocator, WebElement } from 'selenium-webdriver';

const SheetID = process.env.ABEMA_Report_SheetID;
const SheetName = `1週前`;
const StartRow = 2;

// エラー文を格納
const ErrorText = [];
// データの格納
const SheetData = [];

async function Start() {
  // デバッグログを最小限にする
  RPA.Logger.level = 'INFO';
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });
  await GetValue_function();
  await AdjustLOGIN_function();
  for (let i in SheetData) {
    await iOSGetInstall_function(SheetData[i]);
    await AndroidGetInstall_function(SheetData[i]);
    await RPA.Logger.info(SheetData);
  }
  await SetValue_function();
  RPA.Logger.info(ErrorText);
  await RPA.WebBrowser.quit();
}

Start();

// スプレッドシートからデータ取得する関数
async function GetValue_function() {
  try {
    await RPA.Google.authorize({
      //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
    });
    const Data = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: SheetID,
      range: `${SheetName}!B${StartRow}:B8`,
    });
    // 行数を追加する
    for (let i in Data) {
      Data[i].push(Number(i) + StartRow);
      // データを格納する
      SheetData.push(Data[i]);
    }
  } catch (Error) {
    ErrorText[0] = Error;
  }
}

// Adjust ログイン関数
async function AdjustLOGIN_function() {
  try {
    await RPA.WebBrowser.get(process.env.ABEMA_Report_Adjust_URL);
    await RPA.sleep(1000);
    const IDinput = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        id: 'email',
      }),
      15000
    );
    await RPA.WebBrowser.sendKeys(IDinput, [
      process.env.ABEMA_Report_Adjust_ID,
    ]);
    const PWInput = await RPA.WebBrowser.findElementById(`password`);
    await RPA.WebBrowser.sendKeys(PWInput, [
      process.env.ABEMA_Report_Adjust_PW,
    ]);
    const Button = await RPA.WebBrowser.findElementByClassName(`btn btn-block`);
    await RPA.WebBrowser.mouseClick(Button);
    await RPA.sleep(10000);
  } catch (Error) {
    ErrorText[0] = Error;
  }
}

// iOSのインストールページ
async function iOSGetInstall_function(SheetData) {
  try {
    await RPA.WebBrowser.get(
      `${process.env.ABEMA_Report_Adjust_URL_1}${SheetData[0]}&to=${SheetData[0]}${process.env.ABEMA_Report_Adjust_URL_iOS}`
    );
    await RPA.sleep(2000);
    for (let i = 0; i < 5; i++) {
      try {
        const Header = await RPA.WebBrowser.wait(
          RPA.WebBrowser.Until.elementLocated({
            css:
              'body > div > div > div.flex-box-column.flex-one > div > div > div > div > div > div.s-values > div.t-foot.t-scroll > div > div:nth-child(4)',
          }),
          3000
        );
        break;
      } catch {
        await RPA.Logger.info(`【リトライ】${Number(i) + 1}`);
      }
    }
    // ノーデータの表示がある場合は,データ貼り付けに移行する
    const Nodata = await RPA.WebBrowser.findElementsByClassName('no-data');
    if (Nodata.length == 1) {
      await SetValue_function();
    }
    await RPA.sleep(1000);
    const InstallData: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('t-foot t-scroll')[0].children[0].children[3]`
    );
    await SheetData.push(await InstallData.getText());
  } catch (Error) {
    ErrorText[0] = Error;
  }
}

// Androidのインストールページ
async function AndroidGetInstall_function(SheetData) {
  try {
    await RPA.WebBrowser.get(
      `${process.env.ABEMA_Report_Adjust_URL_1}${SheetData[0]}&to=${SheetData[0]}${process.env.ABEMA_Report_Adjust_URL_android}`
    );
    await RPA.sleep(2000);
    for (let i = 0; i < 5; i++) {
      try {
        const Header = await RPA.WebBrowser.wait(
          RPA.WebBrowser.Until.elementLocated({
            css:
              'body > div > div > div.flex-box-column.flex-one > div > div > div > div > div > div.s-values > div.t-foot.t-scroll > div > div:nth-child(4)',
          }),
          3000
        );
        break;
      } catch {
        await RPA.Logger.info(`【リトライ】${Number(i) + 1}`);
      }
    }
    // ノーデータの表示がある場合は,データ貼り付けに移行する
    const Nodata = await RPA.WebBrowser.findElementsByClassName('no-data');
    if (Nodata.length == 1) {
      await SetValue_function();
    }
    await RPA.sleep(1000);
    const InstallData: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('t-foot t-scroll')[0].children[0].children[3]`
    );
    await SheetData.push(await InstallData.getText());
  } catch (Error) {
    ErrorText[0] = Error;
  }
}

// スプレッドシートにデータを貼り付ける関数
async function SetValue_function() {
  await RPA.Logger.info('＊＊＊スプレッドシート貼り付けに移行＊＊＊');
  for (let i in SheetData) {
    // データが4個全て揃っていたら貼り付ける
    if (SheetData[i].length == 4) {
      // iOS
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: SheetID,
        range: `${SheetName}!G${SheetData[i][1]}:G${SheetData[i][1]}`,
        values: [[SheetData[i][2]]],
        parseValues: true,
      });
      // Android
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: SheetID,
        range: `${SheetName}!H${SheetData[i][1]}:H${SheetData[i][1]}`,
        values: [[SheetData[i][3]]],
        parseValues: true,
      });
    }
  }
  await RPA.Logger.info('＊＊＊完了＊＊＊');
  await RPA.WebBrowser.quit();
  await RPA.sleep(2000);
  // 終了
  process.exit(0);
}
