import RPA from 'ts-rpa';
const moment = require('moment');

async function Start() {
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });
  // 実行前にダウンロードフォルダを全て削除する
  await RPA.File.rimraf({ dirPath: `${process.env.WORKSPACE_DIR}` });
  await RPA.Logger.info(RPA.File.outDir);
  // 指定したスプシを pdf で保存
  await GetPDF_function();
  // pdf をドライブにアップロードする
  await DriveUPload_function();
  // 全データシートの転記を行う
  //await Transcribe_function();
  // 古いデータの削除を行う
  await Clear_function();
  // 日付の更新を行う
  await Updata_Date();
  await RPA.WebBrowser.quit();
}

Start();

async function GetPDF_function() {
  // 指定したシートIDをpdfで保存
  await RPA.WebBrowser.get(process.env.ABEMA_Report_pdf_URL);
  await RPA.sleep(3000);
  const IDinput = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      id: 'Email',
    }),
    15000
  );
  await RPA.WebBrowser.sendKeys(IDinput, [`${process.env.ABEMA_ID}`]);
  const NextButton1 = await RPA.WebBrowser.findElementById('next');
  await NextButton1.click();
  await RPA.sleep(3000);
  const PWInput = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      id: 'password',
    }),
    15000
  );
  await RPA.WebBrowser.sendKeys(PWInput, [`${process.env.ABEMA_PW}`]);
  const NextButton2 = await RPA.WebBrowser.findElementById(`submit`);
  await NextButton2.click();
  await RPA.sleep(10000);
}

async function DriveUPload_function() {
  try {
    const FileNames = await RPA.File.list();
    const FileName = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
      range: `全データ!B2:B2`,
    });
    for (let i in FileNames) {
      if (FileNames[i].includes('.pdf') == true) {
        // pdfファイルの名前を日付に変更
        await RPA.File.rename({
          old: `${FileNames[i]}`,
          new: `${FileName[0][0]}週.pdf`,
        });
        break;
      }
    }
    RPA.Logger.info(`＊＊＊ pdfアップロード 開始＊＊＊`);
    await RPA.Google.Drive.upload({
      parents: [`${process.env.ABEMA_Report_pdf_DriveID}`],
      filename: `${FileName[0][0]}週.pdf`,
    });
  } catch (ErrorMes) {
    console.log(ErrorMes);
    await RPA.WebBrowser.quit();
    await RPA.sleep(2000);
    await process.exit();
  }
}

// 全データシートの転記を行う
async function Transcribe_function() {
  RPA.Logger.info(`＊＊＊ 全データ 転記開始 ＊＊＊`);
  // トピック部分の転記
  const Range_C2_E50 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
    range: `全データ!C2:E50`,
  });
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
    range: `全データ!C12:E60`,
    values: Range_C2_E50,
    parseValues: true,
  });

  // iOS/Andrion DL実績/ 新規獲得目線の転記
  const Range_J2_L48 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
    range: `全データ!J2:L48`,
  });
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
    range: `全データ!J12:L58`,
    values: Range_J2_L48,
    parseValues: true,
  });

  // 復帰獲得件数/復帰獲得目線の転記
  const Range_N2_O48 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
    range: `全データ!N2:O48`,
  });
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
    range: `全データ!N12:O58`,
    values: Range_N2_O48,
    parseValues: true,
  });
  RPA.Logger.info(`＊＊＊ 全データ 転記完了 ＊＊＊`);
}

// 転記終わったら 今週分の空白にするためセル削除する
async function Clear_function() {
  RPA.Logger.info(`＊＊＊ 古いデータの削除開始 ＊＊＊`);
  const clear_data = [
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['週間\nトピック1', '週間\nトピック2', '週間\nトピック3'],
    ['', '', ''],
  ];
  // トピックを消す
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
    range: `全データ!C2:E10`,
    values: clear_data,
    parseValues: true,
  });
  // iOS/Androind実績を消す
  const clear_data2 = [
    ['', ''],
    ['', ''],
    ['', ''],
    ['', ''],
    ['', ''],
    ['', ''],
    ['', ''],
  ];
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
    range: `全データ!J2:K8`,
    values: clear_data2,
    parseValues: true,
  });
  // 復帰獲得件数を消す
  const clear_data3 = [[''], [''], [''], [''], [''], [''], ['']];
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
    range: `全データ!N2:N8`,
    values: clear_data3,
    parseValues: true,
  });
  /*
  // 目線を消す
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
    range: `新規/復帰獲得【目線】!D4:E10`,
    values: clear_data2,
    parseValues: true,
  });
  RPA.Logger.info(`＊＊＊ 古いデータの削除完了 ＊＊＊`);
  */
}

// 日付の更新を行う
async function Updata_Date() {
  RPA.Logger.info(`＊＊＊ 日付の更新開始 ＊＊＊`);
  const DayData = moment().format('YYYY-MM-DD');
  console.log(DayData);
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${process.env.ABEMA_Report_SheetID}`,
    range: `全データ!B2:B2`,
    values: [[`${DayData}`]],
    parseValues: true,
  });
  RPA.Logger.info(`＊＊＊ 日付の更新完了 ＊＊＊`);
}
