import RPA from 'ts-rpa';
import { By, WebElement } from 'selenium-webdriver';
const fs = require('fs');
const request = require('request');
const moment = require('moment');
// const ytdl = require('ytdl-core');

// Slack
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [
  `＊＊＊Jenkinsテスト＊＊＊\n【Ameba インスタグラム】完了しました`
];

// スプレッドシートIDとシート名を記載
const mySSID = process.env.My_SheetID;
// const SSID = process.env.Ameba_SheetID2;
const SSName = process.env.Ameba_SheetName;
// 作業するスプレッドシートから読み込む行数を記載
let StartRow = 1;
let LastRow = 30000;

// シートのデータ
let SheetData;
// シート最終行のデータ
let LastRowData;
// 最終行の次の行
let SetDataRow;

// 画像などを保存するフォルダのパスを記載
const DLFolder = __dirname + `/DL`;
// 画像などのURLを格納する変数
let ContentsUrlList = [];
// 画像などのパスを格納する変数
let ContentsPathList = [];

// エラー発生時のテキストを格納
const ErrorText = [];

let Today;
let PostingTime;
let PostFlag = `true`;
let ContentsFlag = `Image`;
async function Start() {
  if (ErrorText.length == 0 && PostFlag == `true` && ContentsFlag == `Image`) {
    // デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
    RPA.Logger.level = 'INFO';
    await RPA.Google.authorize({
      // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    SheetData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${mySSID}`,
      range: `${SSName}!A${StartRow}:Z${LastRow}`
    });
    const SheetTitle = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${mySSID}`,
      range: `${SSName}!B${StartRow}:B${LastRow}`
    });
    LastRow = SheetTitle.length;
    LastRowData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${mySSID}`,
      range: `${SSName}!A${LastRow}:Z${LastRow}`
    });
    await RPA.Logger.info(`シート最終行のデータ　 → ${LastRowData[0]}`);
    await RPA.Logger.info(`前話テーマ　　　　　　 → ${LastRowData[0][9]}`);
    await RPA.Logger.info(`前話リンク　　　　　　 → ${LastRowData[0][10]}`);
    SetDataRow = SheetTitle.length + 1;
    // 本日の日付をフォーマット変更して取得
    Today = moment().format('YYYY/MM/DD');
    // Today = moment()
    //   .add(-7, 'days')
    //   .format('YYYY/MM/DD');
    await RPA.Logger.info(`本日の日付　　　　　→ ` + Today);
    await Work();
    await RPA.Logger.info(`＊＊＊作業を終了します＊＊＊`);
  }
  // エラー発生時の処理
  if (ErrorText.length > 0) {
    // const DOM = await RPA.WebBrowser.driver.getPageSource();
    // await RPA.Logger.info(DOM);
    await RPA.SystemLogger.error(ErrorText);
    Slack_Text[0] = `【Ameba インスタグラム】でエラーが発生しました\n${ErrorText}`;
    // await RPA.WebBrowser.takeScreenshot();
  }
  if (PostFlag == `false`) {
    await RPA.Logger.info(`＊＊＊本日の投稿はありませんでした＊＊＊`);
  }
  if (ContentsFlag == `Video`) {
    await RPA.Logger.info(
      `＊＊＊動画が投稿されていたため、取得をスキップしました＊＊＊`
    );
  }
  await SlackPost(Slack_Text[0]);
  await RPA.WebBrowser.quit();
  await RPA.sleep(1000);
  await process.exit();
}

Start();

async function Work() {
  try {
    // 著名人のインスタに遷移
    // await Instagram();
    // 画像を取得
    // await GetContents();
    // 投稿内容を取得
    // await GetTitle();
    // 取得した画像をダウンロード
    // await ContentsDownload();
    // アメブロにログイン
    // await AmebaLogin();
    // Google Driveへアップロード
    // await DriveUpload();
  } catch (error) {
    ErrorText[0] = error;
    await Start();
  }
}

async function SlackPost(Text) {
  await RPA.Slack.chat.postMessage({
    token: Slack_Token,
    channel: Slack_Channel,
    text: `${Text}`
  });
}

let PostUrl;
async function Instagram() {
  let Element;
  await RPA.WebBrowser.get(process.env.Instagram_HC_URL);
  while (Element == null) {
    await RPA.sleep(2000);
    try {
      Element = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({
          className: ` _2z6nI`
        }),
        15000
      );
      if (Element == null) {
        await RPA.Logger.info(`タイムラインが出ないため、リロードします`);
        await RPA.WebBrowser.refresh();
      } else {
        await RPA.Logger.info(`＊＊＊はあちゅう のInstagramです＊＊＊`);
      }
    } catch {}
  }
  const PostPath: WebElement = await RPA.WebBrowser.driver.executeScript(
    // `return document.getElementsByClassName('v1Nh3 kIKUG  _bz0w')[0].children[0].getAttribute('href')`
    `return document.getElementsByClassName('v1Nh3 kIKUG  _bz0w')[1].children[0].getAttribute('href')`
  );
  await RPA.Logger.info(`投稿ページに直接遷移します`);
  PostUrl = `https://www.instagram.com${PostPath}`;
  await RPA.WebBrowser.get(PostUrl);
  await RPA.sleep(2000);
  const Posting: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('_1o9PC Nzb55')[0].getAttribute('datetime')`
  );
  const PostDate = moment(Posting).format(`YYYY/MM/DD`);
  PostingTime = moment(Posting).format(`YYYY/M/D HH:mm:ss`);
  await RPA.Logger.info(`投稿時点の日付　　　→ ` + PostDate);
  await RPA.Logger.info(`投稿時間　　　　　　→ ` + PostingTime);
  if (Today == PostDate) {
    await RPA.Logger.info(`画像を取得します`);
  } else {
    PostFlag = `false`;
    // E列に記事化しない旨を記載
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${mySSID}`,
      range: `${SSName}!E${SetDataRow}:E${SetDataRow}`,
      values: [[`インスタ投稿なく、記事化なし`]]
    });
    await Start();
  }
}

async function GetContents() {
  let count = 1;
  let ElementFlag = `true`;
  // 画像が2枚以上投稿されているかを確認するため右カーソルの有無を判断
  try {
    const NextContents = await RPA.WebBrowser.findElementsByClassName(
      `    coreSpriteRightChevron  `
    );
    // 右カーソルがある場合（複数）
    if (NextContents.length > 0) {
      await RPA.Logger.info(`右カーソルがあります`);
      while (NextContents.length > 0) {
        await RPA.Logger.info(`${count} 枚目`);
        // 動画か画像か判断
        try {
          if (ElementFlag == `true`) {
            var JudgeContents: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('vi798')[0].children[1].getElementsByTagName('video')[0]`
            );
          } else {
            var JudgeContents: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('vi798')[0].children[2].getElementsByTagName('video')[0]`
            );
          }
          // 動画の場合はスルー
          if (JudgeContents) {
            await RPA.Logger.info(`動画のためスルーします`);
          } else {
            if (ElementFlag == `true`) {
              var Image: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('vi798')[0].children[1].children[0].children[0].children[0].children[0].children[0]`
              );
            }
            // 2枚目以降は"li"タグの2番目を常に取得していく
            else {
              var Image: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('vi798')[0].children[2].children[0].children[0].children[0].children[0].children[0]`
              );
            }
            let ImageUrl;
            ImageUrl = await Image.getAttribute(`src`);
            await RPA.Logger.info(ImageUrl);
            await ContentsUrlList.push(ImageUrl);
            ImageUrl = await ImageUrl.split(/\//);
            let ImagePath;
            ImagePath = await ImageUrl[7].split(`?`);
            ImagePath = await ImagePath[0].split(`.`);
            ImagePath = `${count}.${ImagePath[1]}`;
            await ContentsPathList.push(ImagePath);
          }
        } catch {}
        ElementFlag = `false`;
        await NextContents[0].click();
        await RPA.sleep(1000);
        count++;
      }
    }
    // 右カーソルがない場合（単発）
    else {
      await RPA.Logger.info(`右カーソルはありませんでした`);
      try {
        const JudgeContents: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('_97aPb wKWK0')[0].getElementsByTagName('video')[0]`
        );
        if (JudgeContents) {
          ContentsFlag = `Video`;
          await Start();
        } else {
          const Image: WebElement = await RPA.WebBrowser.driver.executeScript(
            `return document.getElementsByClassName('_97aPb wKWK0')[0].children[0].children[0].children[0].children[0].children[0]`
          );
          let ImageUrl;
          ImageUrl = await Image.getAttribute(`src`);
          await RPA.Logger.info(ImageUrl);
          await ContentsUrlList.push(ImageUrl);
          ImageUrl = await ImageUrl.split(/\//);
          const ImagePath = await ImageUrl[7].split(`?`);
          await RPA.Logger.info(ImagePath[0]);
          await ContentsPathList.push(ImagePath[0]);
        }
      } catch {}
    }
  } catch {}
  if (ContentsUrlList.length < 1) {
    ContentsFlag = `Video`;
    await Start();
  } else {
    await RPA.Logger.info(`画像の取得を終了します`);
  }
  // await RPA.sleep(1000000000);
}

// ※はあちゅうの場合はハッシュタグが固定なので取得しない
// const HashTagList = [];
// const HashTagLinkList = [];
async function GetTitle() {
  // タイトルを取得
  const text: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('C4VMK')[0].children[1]`
  );
  let Title;
  Title = await text.getText();
  Title = await Title.split(`\n`);
  Title = Title[0];
  await RPA.Logger.info(`タイトル　　　　　　→ 【${Title}】`);

  // ※上記理由でコメントアウト
  // const Tag = await text.findElements(By.tagName(`a`));
  // ハッシュタグのリンクを取得
  // let HashTag;
  // for (let i = 0; i <= Tag.length - 1; i++) {
  //   HashTag = await Tag[i].getText();
  //   const HashTagLink = await Tag[i].getAttribute(`href`);
  //   await RPA.Logger.info(`ハッシュタグ　　　　${i + 1} つ目 → ${HashTag}`);
  //   // await RPA.Logger.info(`ハッシュタグリンク　${i + 1} つ目 → ${HashTagLink}`);
  //   if (
  //     HashTag.includes('#') == true &&
  //     HashTagLink.includes('/tags/') == true
  //   ) {
  //     HashTag = await HashTag.slice(1);
  //     await HashTagList.push(HashTag);
  //     await HashTagLinkList.push(HashTagLink);
  //   }
  // }
  // await RPA.Logger.info(HashTagList);

  // B列に取得したURLを記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${mySSID}`,
    range: `${SSName}!B${SetDataRow}:B${SetDataRow}`,
    values: [[PostUrl]]
  });
  // E列に取得したタイトルを記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${mySSID}`,
    range: `${SSName}!E${SetDataRow}:E${SetDataRow}`,
    values: [[Title]]
  });
}

async function ContentsDownload() {
  await RPA.Logger.info(`ダウンロード実行中...`);
  for (let i in ContentsUrlList) {
    await request(
      { method: `GET`, url: ContentsUrlList[i], encoding: null },
      async function(error, response, body) {
        if (!error && response.statusCode === 200) {
          await fs.writeFileSync(
            `${DLFolder}/${ContentsPathList[i]}`,
            body,
            `binary`
          );
        }
      }
    );
  }
  await RPA.Logger.info(`ダウンロード完了しました`);
}

let LoadingFlag = `false`;
async function AmebaLogin() {
  // try {
  await RPA.Logger.info(`アベブロにログインします`);
  await RPA.WebBrowser.get(process.env.Instagram_Login_URL);
  await RPA.sleep(2000);
  const LoginID = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ name: `accountId` }),
    15000
  );
  await RPA.WebBrowser.sendKeys(LoginID, [process.env.AmebaLogin_ID_YY]);
  await RPA.sleep(500);
  const LoginPW = await RPA.WebBrowser.driver.findElement(By.name(`password`));
  await RPA.WebBrowser.sendKeys(LoginPW, [process.env.AmebaLogin_PW_YY]);
  await RPA.sleep(500);
  const LoginButton = await RPA.WebBrowser.findElementByClassName(
    `c-btn c-btn--large c-btn--primary`
  );
  await LoginButton.click();
  await RPA.sleep(3000);
  // const Alert = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({
  //     className: `sqdOP yWX7d    y3zKF     `
  //   }),
  //   15000
  // );
  // await Alert.click();
  // await RPA.sleep(1000);
  // const Alert2 = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({ className: `aOOlW   HoLwm ` }),
  //   15000
  // );
  // await Alert2.click();
  // while (LoadingFlag == `false`) {
  //   await RPA.sleep(2000);
  //   try {
  //     const Start = await RPA.WebBrowser.wait(
  //       RPA.WebBrowser.Until.elementLocated({
  //         className: `_7UhW9   xLCgt       qyrsm KV-D4         uL8Hv         `
  //       }),
  //       15000
  //     );
  //     const StartText = await Start.getText();
  //     if (StartText.length < 1) {
  //       await RPA.Logger.info(`タイムラインが出ないため、リロードします`);
  //       await RPA.WebBrowser.refresh();
  //       await RPA.sleep(1000);
  //     } else {
  //       LoadingFlag = `true`;
  //       await RPA.Logger.info(StartText);
  //       await RPA.Logger.info(`＊＊＊ログイン成功しました＊＊＊`);
  //     }
  //   } catch {}
  // }
  //   LoadingFlag = `false`;
  // } catch (error) {
  //   Login = `Failure`;
  //   Slack_Text[0] = `【Ameba インスタグラム】ログインに失敗しました！\n${error}`;
  //   await Start();
  // }
  // そのままでは絵文字が入力できないためエンコードする
  // const Text1 = await encodeURI(Title);
  // await RPA.WebBrowser.driver.executeScript(
  //   `document.getElementsByClassName('KYqgv_Qj--TweetComposer-tweetTextarea')[0].value = decodeURI('${Text1}')`
  // );
  // await RPA.WebBrowser.sendKeys(PostWording, [` `]);
  // // 前話リンクを作成
  // const ZenwaLink = await RPA.WebBrowser.driver.findElement(
  //   By.name(`password`)
  // );
  // let ZenwaLinkText = await ZenwaLink.getText();
  // // 一旦入力してあるものを取得
  // await RPA.Logger.info(ZenwaLinkText);
  // // ※それをコピペしてこっちで作成
  // ZenwaLinkText = `${LastRowData[0][9]}`;
  // await ZenwaLinkText.clear();
  // await RPA.sleep(100);
  // await RPA.WebBrowser.sendKeys(ZenwaLinkText, [LastRowData[0][9]]);
  // await RPA.sleep(300);
  // const ZenwaLinkUrl = await RPA.WebBrowser.driver.findElement(
  //   By.name(`password`)
  // );
  // await ZenwaLinkText.clear();
  // await RPA.sleep(100);
  // await RPA.WebBrowser.sendKeys(ZenwaLinkUrl, [LastRowData[0][10]]);
  // await RPA.sleep(300);
  // ※最後に必ずログアウト処理を入れること（入れなくても良いかも）
}

async function DriveUpload() {
  await RPA.Logger.info(`アップロード実行中...`);
  for (let i in ContentsPathList) {
    // await RPA.sleep(3000);
    await RPA.Google.Drive.upload({
      filename: ContentsPathList[i],
      parents: [`${process.env.My_GoogleDriveFolderID2}`]
    });
  }
  await RPA.Logger.info(`アップロード完了しました`);
}
