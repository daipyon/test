import RPA from 'ts-rpa';
import { By, WebElement } from 'selenium-webdriver';
const fs = require('fs');
const moment = require('moment');

const Today = moment().format('YYYY-MM-DD');
RPA.Logger.info(Today);

// 本番用
// const Slack_Token = process.env.ABEMA_Hollywood_bot_token;
// const Slack_Channel = process.env.RPA_Test_Channel;

// テスト用
const Slack_Token = process.env.Test_Slack_Token;
const Slack_Channel = process.env.Test_Slack_Channel;

let Slack_Text = ``;

const Tableau_URL9 = process.env.Spot_Tablaeu_Login_URL9;
const CsvDownloader_URL = process.env.Spot_CsvDownloader_URL;
const WorkingName = '投げ銭戦略カレンダー';

async function Start() {
  try {
    // デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
    // RPA.Logger.level = 'INFO';
    // 実行前にダウンロードフォルダを全て削除する
    await RPA.File.rimraf({ dirPath: `${process.env.WORKSPACE_DIR}` });
    await CASSO_LOGIN_function();
    await Download_function();
    await Rename_function();
    await SlackFilePost_function(Slack_Text);
    await RPA.sleep(2000);
    RPA.Logger.info(`【ハリウッド】${WorkingName}完了しました`);
  } catch (error) {
    // const DOM = await RPA.WebBrowser.driver.getPageSource();
    // await RPA.Logger.info(DOM);
    await RPA.SystemLogger.error(error);
    Slack_Text = `【ハリウッド】でエラーが発生しました\n${error}`;
    await RPA.WebBrowser.takeScreenshot();
    await SlackFilePost_function(Slack_Text);
  }
  await RPA.WebBrowser.quit();
  await RPA.sleep(1000);
  await process.exit();
}

Start();

async function CASSO_LOGIN_function() {
  await RPA.WebBrowser.get(
    'https://tableau.cyberagent.group/#/site/AbemaTV/views/_43/sheet1?:iid=1'
  );
  const idinput = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'username' }),
    15000
  );
  await RPA.WebBrowser.sendKeys(idinput, [process.env.CyNumber]);
  const PW_input = await RPA.WebBrowser.findElementById(`password`);
  await RPA.WebBrowser.sendKeys(PW_input, [process.env.CyPass]);
  const NextButton = await RPA.WebBrowser.findElementByCSSSelector(
    `body > div > div.ping-body-container > div > form > div.ping-buttons > a`
  );
  await RPA.WebBrowser.mouseClick(NextButton);
  await RPA.sleep(2000);
  // タブロー操作用のフレームに切り替え
  const Ifream = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      css: '#viz > iframe'
    }),
    10000
  );
  await RPA.WebBrowser.switchToFrame(Ifream);
  await RPA.sleep(5000);
}

let StartDate;
let EndDate;

async function Download_function() {
  RPA.Logger.info(`CSVダウンロード開始します`);
  const DLButton = await RPA.WebBrowser.findElementById(
    `download-ToolbarButton`
  );
  await DLButton.click();
  await RPA.sleep(2000);
  await RPA.WebBrowser.driver.executeScript(
    `document.getElementById('DownloadDialog-Dialog-Body-Id').children[0].children[3].click()`
  );
  await RPA.sleep(4000);
  // const BUTTON: WebElement =
  await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('fdiufnn low-density')[0].click()`
  );
  // const BUTTON = await RPA.WebBrowser.findElementByXPath(
  //   `//*[@id="export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id"]/div/div[2]/button`
  // );
  // RPA.Logger.log(BUTTON);
  // await BUTTON.click();
  await RPA.sleep(10000);
  RPA.Logger.info(`【タブロー】ダウンロード完了`);
}

async function Rename_function() {
  const FileList = await RPA.File.list();
  RPA.Logger.info(FileList);
  for (let i in FileList) {
    if (FileList[i].includes('.csv') == true) {
      await RPA.File.rename({
        old: FileList[i],
        new: `${WorkingName}${Today}.csv`
      });
      RPA.Logger.info('【CSV】リネーム完了');
      break;
    }
  }
}

async function SlackFilePost_function(Slack_Text) {
  await RPA.Slack.files.upload({
    token: Slack_Token,
    // s が付いていないと効かない
    // channel: Slack_Channel,
    channels: Slack_Channel,
    text: `${Slack_Text}`,
    file: fs.createReadStream(__dirname + `/DL/${WorkingName}${Today}.csv`)
  });
}
