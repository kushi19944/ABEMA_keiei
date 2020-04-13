import RPA from 'ts-rpa';

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
  await GoogleLogIN();
  // pdf をドライブにアップロードする
  await DriveUPload();
  await RPA.WebBrowser.quit();
}

Start();

async function GoogleLogIN() {
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

async function DriveUPload() {
  const FileNames = await RPA.File.list();
  var today = new Date();
  const ToDay =
    today.getFullYear() + `-` + (today.getMonth() + 1) + `-` + today.getDate();
  for (let i in FileNames) {
    if (FileNames[i].includes('.pdf') == true) {
      // pdfファイルの名前を日付に変更
      await RPA.File.rename({ old: `${FileNames[i]}`, new: `${ToDay}.pdf` });
      break;
    }
  }
  RPA.Logger.info(`＊＊＊ pdfアップロード 開始＊＊＊`);
  await RPA.Google.Drive.upload({
    parents: [`${process.env.ABEMA_Report_pdf_DriveID}`],
    filename: `${ToDay}.pdf`,
  });
}
