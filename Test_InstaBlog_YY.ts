import RPA from 'ts-rpa';
import { By, WebElement } from 'selenium-webdriver';
const fs = require('fs');
const request = require('request');
const moment = require('moment');

// Slack
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
// const Slack_Text = [`【Ameba インスタグラム】完了しました`];
const Slack_Text = [`＊＊＊テスト＊＊＊【Ameba インスタグラム】完了しました`];

// 画像などを保存するフォルダのパスを記載
const DLFolder = __dirname + `/DL`;

// 画像などのURLを格納する変数
let ContentsUrlList = [];
// 画像などのパスを格納する変数
let ContentsPathList = [];

// エラー発生時のテキストを格納
const ErrorText = [];

let Today;
let Posting;
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
    // 本日の日付をフォーマット変更して取得
    // Today = moment().format('YYYY/MM/DD');
    Today = moment()
      .add(-28, 'days')
      .format('YYYY/MM/DD');
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
    await RPA.WebBrowser.takeScreenshot();
  }
  if (PostFlag == `false`) {
    await RPA.Logger.info(`＊＊＊本日の投稿はありませんでした＊＊＊`);
  }
  if (ContentsFlag == `Video`) {
    await RPA.Logger.info(
      `＊＊＊動画が投稿されていたため、更新作業を終了します＊＊＊`
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
    await Instagram();
    // 画像を取得
    await GetContents();
    // 投稿内容を取得
    await GetText();
    // 取得した画像をダウンロード
    await ContentsDownload();
    // アメブロにログイン
    await AmebaLogin();
    // 画像・投稿内容などを入力
    await InputData();
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

let AccessFlag = `Bad`;
let Element;
async function Instagram() {
  // await RPA.sleep(2000);
  await RPA.WebBrowser.get(process.env.Instagram_YY_URL);
  // while (AccessFlag == `Bad`) {
  await RPA.sleep(2000);
  //   try {
  //     const CantAccess = await RPA.WebBrowser.findElementById(`main-message`);
  //     const CantAccessMessage = await CantAccess.getText();
  //     if (CantAccessMessage == `このサイトにアクセスできません`) {
  //       await RPA.Logger.info(`アクセスできないため、リロードします`);
  //       await RPA.WebBrowser.refresh();
  //     } else {
  //       AccessFlag = `Good`;
  await RPA.Logger.info(`＊＊＊山田優 のInstagramです＊＊＊`);
  //     }
  //   } catch {}
  // }
  // document.getElementById(`main-message`)
  // a.children[0].children[0]
  const PostUrl: WebElement = await RPA.WebBrowser.driver.executeScript(
    // `return document.getElementsByClassName('v1Nh3 kIKUG  _bz0w')[0].children[0].getAttribute('href')`
    `return document.getElementsByClassName('v1Nh3 kIKUG  _bz0w')[8].children[0].getAttribute('href')`
  );
  await RPA.Logger.info(`投稿ページに直接遷移します`);
  await RPA.WebBrowser.get(`https://www.instagram.com${PostUrl}`);
  await RPA.sleep(2000);
  const DateTime: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('_1o9PC Nzb55')[0].getAttribute('datetime')`
  );
  Posting = DateTime;
  const PostDate = moment(Posting).format(`YYYY/MM/DD`);
  PostingTime = moment(Posting).format(`YYYY/M/D HH:mm:ss`);
  await RPA.Logger.info(`投稿時点の日付　　　→ ` + PostDate);
  await RPA.Logger.info(`投稿時間　　　　　　→ ` + PostingTime);
  if (Today == PostDate) {
    await RPA.Logger.info(`画像を取得します`);
  } else {
    PostFlag = `false`;
    await Start();
  }
}

let count = 1;
let ElementFlag = `true`;
async function GetContents() {
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
                `return document.getElementsByClassName('vi798')[0].getElementsByTagName('img')[0]`
              );
            }
            // 2枚目以降は"li"タグの2番目を常に取得していく
            else {
              var Image: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('vi798')[0].getElementsByTagName('img')[1]`
              );
            }
            const ImageUrl = await Image.getAttribute(`src`);
            await RPA.Logger.info(ImageUrl);
            await ContentsUrlList.push(ImageUrl);
            const ImagePath = `${count}.jpg`;
            // await RPA.Logger.info(ImagePath);
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
            `return document.getElementsByClassName('_97aPb wKWK0')[0].getElementsByTagName('img')[0]`
          );
          const ImageUrl = await Image.getAttribute(`src`);
          await RPA.Logger.info(ImageUrl);
          await ContentsUrlList.push(ImageUrl);
          const ImagePath = `1.jpg`;
          await ContentsPathList.push(ImagePath);
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
}

let PostContent;
const HashTagList = [];
async function GetText() {
  const InnerHtml: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('C4VMK')[0].children[1].innerText`
  );
  PostContent = InnerHtml;
  await RPA.Logger.info(PostContent);
  // ハッシュタグのリンクを取得
  const text: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('C4VMK')[0].children[1]`
  );
  const Tag = await text.findElements(By.tagName(`a`));
  let HashTag;
  for (let i = 0; i <= Tag.length - 1; i++) {
    HashTag = await Tag[i].getText();
    const HashTagLink = await Tag[i].getAttribute(`href`);
    // await RPA.Logger.info(`ハッシュタグ　　　　${i + 1} つ目 → ${HashTag}`);
    if (
      HashTag.includes('#') == true &&
      HashTagLink.includes('/tags/') == true
    ) {
      await HashTagList.push(HashTag);
      HashTag = await HashTag.slice(1);
    }
  }
  await RPA.Logger.info(HashTagList);
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
  await RPA.Logger.info(`＊＊＊ダウンロード完了しました＊＊＊`);
}

async function AmebaLogin() {
  await RPA.Logger.info(`アベブロにログインします`);
  await RPA.WebBrowser.get(process.env.Ameba_Login_URL);
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
  while (Element == null) {
    await RPA.sleep(2000);
    try {
      Element = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({
          className: `ucs-modal__main`
        }),
        15000
      );
      if (Element == null) {
        await RPA.Logger.info(`編集画面が表示されないため、リロードします`);
        await RPA.WebBrowser.refresh();
      } else {
        await RPA.Logger.info(`＊＊＊ログイン成功しました＊＊＊`);
      }
    } catch {}
  }
  const CloseButton: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('ucs-modal__closeIcon s s-close')[1]`
  );
  await CloseButton.click();
  await RPA.sleep(3000);
}

async function InputData() {
  // タイトルを入れないと下書き保存できないため、適当に文字を入れておく
  const EntryTitle = await RPA.WebBrowser.driver.findElement(
    By.name(`entry_title`)
  );
  await RPA.WebBrowser.sendKeys(EntryTitle, [`下書き保存完了しました`]);
  await RPA.sleep(1000);
  let iframe;
  let element;
  for (let i in ContentsPathList) {
    // InputFile に直接パスを打ち込んでアップロード
    const InputFile = await RPA.WebBrowser.driver.findElement(
      By.name(`thumbnail`)
    );
    await RPA.WebBrowser.sendKeys(InputFile, [
      `${DLFolder}/${ContentsPathList[i]}`
    ]);
    await RPA.Logger.info(`${Number(i) + 1}枚目 アップロード完了しました`);
    await RPA.sleep(3000);
    // アップロードした画像をクリック
    const UploadedImage: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('p-images-imageList__listItem')[1]`
    );
    await UploadedImage.click();
    await RPA.sleep(1000);
    // iframe の中に入る
    iframe = await RPA.WebBrowser.findElementByClassName(
      `cke_wysiwyg_frame cke_reset`
    );
    await RPA.WebBrowser.switchToFrame(iframe);
    element = await RPA.WebBrowser.findElementByClassName(
      `cke_editable cke_editable_themed cke_contents_ltr`
    );
    // エンターキーを二回押して改行
    await RPA.WebBrowser.sendKeys(element, [RPA.WebBrowser.Key.ENTER]);
    await RPA.WebBrowser.sendKeys(element, [RPA.WebBrowser.Key.ENTER]);
    if (Number(i) == ContentsPathList.length - 1) {
      await RPA.WebBrowser.sendKeys(element, [RPA.WebBrowser.Key.BACK_SPACE]);
    }
    // iframeから抜け出す
    await RPA.WebBrowser.driver.switchTo().defaultContent();
  }
  await RPA.sleep(1000);
  await RPA.WebBrowser.switchToFrame(iframe);
  // 投稿内容を入力
  // const INPUT_EMOJI = `arguments[0].value += arguments[1];arguments[0].dispatchEvent(new Event('change'));`;
  const InputEmoji = `arguments[0].innerText += arguments[1];arguments[0].dispatchEvent(new Event('change'));`;
  const n = ContentsPathList.length * 2 - 1;
  const element2: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('cke_editable cke_editable_themed cke_contents_ltr')[0].children[${n}]`
  );
  await RPA.WebBrowser.driver.executeScript(InputEmoji, element2, PostContent);
  await RPA.sleep(1000);
  await RPA.WebBrowser.driver.switchTo().defaultContent();
  // ハッシュタグを入力
  const HashTagEdit = await RPA.WebBrowser.findElementById(`js-hashtag-tags`);
  await HashTagEdit.click();
  await RPA.sleep(1000);
  for (let i in HashTagList) {
    const HashTagInput: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementById('js-hashtag-user-tags').children[0].children[${i}].children[0].children[0]`
    );
    await RPA.WebBrowser.sendKeys(HashTagInput, [HashTagList[i]]);
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(HashTagInput, [RPA.WebBrowser.Key.ENTER]);
    await RPA.sleep(1000);
  }
  const OkButton = await RPA.WebBrowser.findElementById(`js-hashtag-fixButton`);
  await OkButton.click();
  // 投稿時間を入力
  const PostDateResult = await RPA.WebBrowser.findElementById(
    `js-postDateResult`
  );
  await PostDateResult.click();
  await RPA.sleep(1000);
  const PostYear = moment(Posting).format(`YYYY`);
  const PostMonth = moment(Posting).format(`MM`);
  const PostDay = moment(Posting).format(`DD`);
  const PostHours = moment(Posting).format(`HH`);
  const PostMinutes = moment(Posting).format(`mm`);
  const PostSeconds = moment(Posting).format(`ss`);
  const InputYear = await RPA.WebBrowser.findElementById(`js-calInputYear`);
  await RPA.WebBrowser.sendKeys(InputYear, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.WebBrowser.sendKeys(InputYear, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.WebBrowser.sendKeys(InputYear, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.WebBrowser.sendKeys(InputYear, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(InputYear, [PostYear]);
  await RPA.sleep(1000);
  const InputMonth = await RPA.WebBrowser.findElementById(`js-calInputMonth`);
  await RPA.WebBrowser.sendKeys(InputMonth, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.WebBrowser.sendKeys(InputMonth, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(InputMonth, [PostMonth]);
  await RPA.sleep(1000);
  const InputDay = await RPA.WebBrowser.findElementById(`js-calInputDay`);
  await RPA.WebBrowser.sendKeys(InputDay, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.WebBrowser.sendKeys(InputDay, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(InputDay, [PostDay]);
  await RPA.sleep(1000);
  const InputHours = await RPA.WebBrowser.findElementById(`js-calInputHours`);
  await RPA.WebBrowser.sendKeys(InputHours, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.WebBrowser.sendKeys(InputHours, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(InputHours, [PostHours]);
  await RPA.sleep(1000);
  const InputMinutes = await RPA.WebBrowser.findElementById(
    `js-calInputMinutes`
  );
  await RPA.WebBrowser.sendKeys(InputMinutes, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.WebBrowser.sendKeys(InputMinutes, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(InputMinutes, [PostMinutes]);
  await RPA.sleep(1000);
  const InputSeconds = await RPA.WebBrowser.findElementById(
    `js-calInputSeconds`
  );
  await RPA.WebBrowser.sendKeys(InputSeconds, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.WebBrowser.sendKeys(InputSeconds, [RPA.WebBrowser.Key.BACK_SPACE]);
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(InputSeconds, [PostSeconds]);
  await RPA.sleep(1000);
  const OkButton2 = await RPA.WebBrowser.findElementById(`js-calOkButton`);
  await OkButton2.click();
  await RPA.sleep(1000);
  // await RPA.sleep(1000000000);

  // 本番用
  // 下書き保存をクリック
  // const SaveDraft = await RPA.WebBrowser.findElementByClassName(
  //   `js-submitButton c-button c-button--normal`
  // );
  // await SaveDraft.click();
  // await RPA.Logger.info(`＊＊＊下書き保存が完了しました＊＊＊`);
  await RPA.Logger.info(`＊＊＊下書き保存が完了したと推定＊＊＊`);
  await RPA.sleep(5000);
}
