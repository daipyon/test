import RPA from 'ts-rpa';
import { By, Key, WebElement } from 'selenium-webdriver';
// const fs = require('fs');

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Twitter 予約投稿】予約完了しました`];

// スプレッドシートIDとシート名を記載
const mySSID = process.env.My_SheetID;
const SSID = process.env.Senden_Twitter_SheetID;
const SSName = process.env.Senden_Twitter_SheetName;
// アイパス一覧シート
const SSID2 = process.env.Senden_Twitter_SheetID2;
const SSName2 = process.env.Senden_Twitter_SheetName2;

// 画像のパスを保持する変数
const FilePathData = [''];
// パスに記載されている日付を保持する変数
const FilePathDate = [''];
// 作業対象行とデータを取得
const WorkData = [];
const Row = [];

// 画像などを保存するフォルダのパスを記載 ※.envファイルは同じにしない
// const myRPAFolder = process.env.My_RPAFolder;
const RPAFolder = process.env.Senden_Twitter_RPAFolder;

// エラー発生時のテキストを格納
const ErrorText = [];

async function Start() {
  if (ErrorText.length == 0) {
    // デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
    RPA.Logger.level = 'INFO';
    await RPA.Google.authorize({
      // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    // ダウンロードフォルダから画像のパスを取得
    const path = require('path');
    // const dirPath = path.resolve(myRPAFolder);
    const FirstData = [];
    // FirstData[0] = fs.readdirSync(dirPath);
    for (let i in FirstData[0]) {
      if (FirstData[0][i].indexOf('.png') > 0) {
        FilePathData[0] = FirstData[0][i];
        await RPA.Logger.info(FilePathData);
        await Work();
      }
    }
  }
  // エラー発生時の処理
  if (ErrorText.length >= 1) {
    // const DOM = await RPA.WebBrowser.driver.getPageSource();
    // await await RPA.Logger.info(DOM);
    await RPA.SystemLogger.error(ErrorText);
    Slack_Text[0] = `【Twitter 予約投稿】でエラー発生しました\n${ErrorText}`;
    await RPA.WebBrowser.takeScreenshot();
  }
  await RPA.Logger.info('作業を終了します');
  // await SlackPost(Slack_Text[0]);
  await RPA.WebBrowser.quit();
  await RPA.sleep(1000);
  await process.exit();
}

Start();

const FirstLoginFlag = ['true'];
async function Work() {
  try {
    await GetDataRow(FilePathData, FilePathDate, WorkData, Row);
    // Twitterにログイン
    if (FirstLoginFlag[0] == 'true') {
      await TwitterLogin();
    }
    if (FirstLoginFlag[0] == 'false') {
      await TwitterLogin2();
    }
    // 一度ログインしたら、以降はログインページをスキップ
    FirstLoginFlag[0] = 'false';
    // ウェブサイトカードを作成
    await CreateWebSiteCard(FilePathData, FilePathDate);
    // 作成したカード名を検索
    // await SearchCardName(FilePathDate);
    // Twitter投稿を予約
    // await ReservationTweet(WorkData, FilePathData);
    // 作業用フォルダを空にするためファイルを削除
    // await DeleteFile();
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

async function GetDataRow(FilePathData, FilePathDate, WorkData, Row) {
  // ".png"の文字を削除
  FilePathDate[0] = await FilePathData[0].slice(0, -4);
  await RPA.Logger.info(FilePathDate);
  const JudgeData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${mySSID}`,
    range: `${SSName}!C6:C30000`
  });
  for (let i in JudgeData) {
    if (FilePathDate[0] == JudgeData[i][0]) {
      await RPA.Logger.info(JudgeData[i][0]);
      Row[0] = Number(i) + 6;
      break;
    }
  }
  await RPA.Logger.info('この行のデータを取得します → ', Row[0]);
  // 番組表下書きシートのデータ(B〜AA列)を取得
  WorkData[0] = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${mySSID}`,
    range: `${SSName}!B${Row[0]}:AA${Row[0]}`
  });
  await RPA.Logger.info(WorkData[0]);
}

async function TwitterLogin() {
  // throw new Error();
  // アイパス一覧シートからTwitterのIDとパスワードを取得
  const TwitterID = await RPA.Google.Spreadsheet.getValues({
    // spreadsheetId: `${SSID2}`,
    spreadsheetId: `${mySSID}`,
    range: `${SSName2}!D3:D3`
  });
  const TwitterPW = await RPA.Google.Spreadsheet.getValues({
    // spreadsheetId: `${SSID2}`,
    spreadsheetId: `${mySSID}`,
    range: `${SSName2}!E3:E3`
  });
  await RPA.Logger.info(TwitterID);
  await RPA.Logger.info(TwitterPW);
  // Twitterにログイン
  await RPA.WebBrowser.get(process.env.Twitter_Login_URL);
  await RPA.sleep(5000);
  const Login: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('css-1dbjc4n r-1awozwy r-1pz39u2 r-18u37iz r-16y2uox')[0].children[0].children[0]`
  );
  await RPA.WebBrowser.mouseClick(Login);
  await RPA.sleep(1000);
  const TwitterLoginId = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      name: 'session[username_or_email]'
    }),
    5000
  );
  await RPA.WebBrowser.sendKeys(TwitterLoginId, [TwitterID[0][0]]);
  await RPA.sleep(500);
  const TwitterLoginPw = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ name: 'session[password]' }),
    5000
  );
  await RPA.WebBrowser.sendKeys(TwitterLoginPw, [TwitterPW[0][0]]);
  const LoginButton: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('css-1dbjc4n')[21].children[0]`
  );
  await RPA.WebBrowser.mouseClick(LoginButton);
  while (0 == 0) {
    await RPA.sleep(5000);
    const AbeaniTweet = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({ id: 'accessible-list-0' }),
      5000
    );
    const AbeaniTweetText = await AbeaniTweet.getText();
    if (AbeaniTweetText.length >= 1) {
      await RPA.Logger.info('＊＊＊ログイン成功しました＊＊＊');
      ErrorText.length = 0;
      break;
    }
  }
  await RPA.Logger.info('カードライブラリページに直接遷移します');
  await RPA.WebBrowser.get(process.env.Twitter_CardLibrary_URL);
  await RPA.sleep(5000);
}

async function TwitterLogin2() {
  await RPA.Logger.info('カードライブラリページに直接遷移します');
  await RPA.WebBrowser.get(process.env.Twitter_CardLibrary_URL);
  await RPA.sleep(5000);
}

async function CreateWebSiteCard(FilePathData, FilePathDate) {
  // 【カードを作成】をクリック
  const CreateCard = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      xpath: '//*[@id="root"]/div/div[1]/div[1]/div/button'
    }),
    5000
  );
  await RPA.WebBrowser.mouseClick(CreateCard);
  // 【ウェブサイトカード】をクリック
  const WebSiteCard = await RPA.WebBrowser.findElementById(
    'feather-dropdown-0-menu-item-content-0'
  );
  await RPA.WebBrowser.mouseClick(WebSiteCard);
  // // 【画像を選択】をクリック
  // const SelectImage: WebElement = await RPA.WebBrowser.driver.executeScript(
  //   `return document.getElementsByClassName('FormField')[0].children[1]`
  // );
  // await RPA.WebBrowser.mouseClick(SelectImage);
  // // InputFile に直接パスを打ち込んでアップロード
  // const InputFile = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({
  //     xpath: '/html/body/div[5]/div/div/div[2]/div[3]/div/div[2]/div[1]/input'
  //   }),
  //   5000
  // );
  // await RPA.WebBrowser.sendKeys(InputFile, [
  //   `${myRPAFolder}/${FilePathData[0]}`
  // ]);
  // await RPA.sleep(3000);
  // 【ヘッドライン】を入力
  const HeadLine: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('FormInput')[1]`
  );
  await RPA.WebBrowser.sendKeys(HeadLine, ['△タップして番組表一覧を見る']);
  // 【ウェブサイトのURL】を入力
  const WebSiteUrl: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('FormInput')[2]`
  );
  await RPA.WebBrowser.sendKeys(WebSiteUrl, [process.env.Twitter_WebSite_URL]);
  // 【カード名】を入力
  const CardName: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('FormInput')[3]`
  );

  // テスト用
  await RPA.WebBrowser.sendKeys(CardName, [
    `※※※RPAテストです※※※ ${FilePathDate[0]}`
  ]);

  // 本番用
  // await RPA.WebBrowser.sendKeys(CardName, [FilePathDate[0]]);

  // // 【作成】をクリック
  // const Create: WebElement = await RPA.WebBrowser.driver.executeScript(
  //   `return document.getElementsByClassName('Button Button--primary Button--small is-disabled')[0]`
  // );
  // await RPA.WebBrowser.mouseClick(Create);
  await RPA.sleep(5000);
  await RPA.WebBrowser.takeScreenshot();
}

const LoadingFlag = ['false'];
async function SearchCardName(FilePathDate) {
  while (0 == 0) {
    try {
      // 読み込みエラーが発生した場合はブラウザを更新
      await BrowserReload();
    } catch {
      LoadingFlag[0] = 'true';
      await RPA.Logger.info('読み込みエラーは発生しませんでした。');
      break;
    }
  }
  // 再度読み込みエラーが発生した場合は、ブラウザを更新し再検索する
  while (0 == 0) {
    // カード名を検索
    for (let i = 0; i <= 9; i++) {
      const SearchCardName: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('src-cardsmanager-views-CardsGrid-styles-module--cardName')[${Number(
          i
        )}]`
      );
      const SearchCardNameText = await SearchCardName.getText();
      await RPA.Logger.info(SearchCardNameText);

      // テスト用
      if (SearchCardNameText == `※※※RPAテストです※※※ ${FilePathDate[0]}`) {
        // 本番用
        // if (SearchCardNameText == FilePathDate[0]) {

        // 作成したカードにマウスオーバー
        await RPA.Logger.info('    ↑     一致しました');
        const CreatedCard: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('src-cardsmanager-views-CardsGrid-styles-module--cardName')[${Number(
            i
          )}].children[1]`
        );
        await RPA.WebBrowser.mouseMove(CreatedCard);
        // 【ツイート】をクリック
        await RPA.WebBrowser.mouseClick(CreatedCard);
        break;
      }
    }
    try {
      LoadingFlag[0] = 'false';
      await BrowserReload();
    } catch {
      LoadingFlag[0] = 'true';
      await RPA.Logger.info('読み込みエラーは発生しませんでした。');
      break;
    }
  }
  await RPA.sleep(5000);
}

async function ReservationTweet(WorkData, FilePathData) {
  // ウィンドウの切り替え
  const windows = await RPA.WebBrowser.getAllWindowHandles();
  await RPA.WebBrowser.switchToWindow(windows[1]);
  // M列（文言）を入力
  await RPA.Logger.info(WorkData[0][0][11]);
  const PostWording = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      xpath:
        '//*[@id="root"]/div/div[1]/div/div/div/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div[1]/textarea'
    }),
    5000
  );
  // そのままでは絵文字が入力できないためエンコードする
  const Text1 = await encodeURI(WorkData[0][0][11]);
  await RPA.WebBrowser.driver.executeScript(
    `document.getElementsByClassName("KYqgv_Qj--TweetComposer-tweetTextarea")[0].value = decodeURI("${Text1}")`
  );
  await RPA.WebBrowser.sendKeys(PostWording, [` `]);
  await RPA.WebBrowser.scrollTo({
    selector:
      '#root > div > div.src-tweetcomposer-components-PageLayout-styles-module--root > div > div > div > div > div > div.src-tweetcomposer-components-TweetComposerContainer-styles-module--tweetComposerContainer > div > div > div._1hGMbSMU--TweetComposer-actionsContainer'
  });
  // 【広告用】のチェックを外す
  const Uncheck: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('_227h4M8t--FanoutSwitch-checkbox')[0]`
  );
  await RPA.WebBrowser.mouseClick(Uncheck);
  // カーソルをクリック
  const Cursor: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('ButtonGroup ButtonGroup--primary w0xuLYbb--TweetComposer-submitTweetButton')[0].children[1]`
  );
  await RPA.WebBrowser.mouseClick(Cursor);
  await RPA.sleep(1000);
  // 【予約設定する】をクリック
  const ReservationSettings: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('Dropdown-menuItemList')[0].children[0].children[0]`
  );
  await RPA.WebBrowser.mouseClick(ReservationSettings);
  // カレンダーをクリック
  const Calender: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('Button _1xl1L2Qp--TweetComposer-actionButton DatePickerDropdownTarget')[0]`
  );
  await RPA.WebBrowser.mouseClick(Calender);
  // カレンダーの月を取得
  const MonthButton: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('Button Button--tertiary Button--xsmall')[0]`
  );
  const MonthText = await MonthButton.getText();

  // テスト用

  // 本番用
  // カレンダーの"月"という文字を削除
  const Month = await MonthText.slice(0, -1);
  await RPA.Logger.info(Month);
  // ファイル名の月の部分を取得
  const FilePath_Month = await FilePathData[0].slice(5, -6);
  await RPA.Logger.info(FilePath_Month);
  if (Number(Month) != Number(FilePath_Month)) {
    await RPA.Logger.info('月が不一致のため、次月を指定します');
    const NextMonthButton = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="feather-dropdown-7"]/ul/li/div/div[1]/button[2]'
    );
    await RPA.WebBrowser.mouseClick(NextMonthButton);
  }
  if (Number(Month) == Number(FilePath_Month)) {
    await RPA.Logger.info('月が一致しました');
  }
  // ファイル名の日付の部分を取得
  const FilePath_Day = await FilePathData[0].slice(7, -4);
  await RPA.Logger.info(FilePath_Day);
  // 1日の位置を検索
  const ThroughDay = [];
  for (let i = 1; i <= 7; i++) {
    const SearchDay = await RPA.WebBrowser.findElementByXPath(
      `//*[@id="feather-dropdown-7"]/ul/li/div/div[2]/div[2]/div[1]/div[${Number(
        i
      )}]/span`
    );
    const SearchDayText = await SearchDay.getText();
    if (SearchDayText != '1') {
      await RPA.Logger.info(`${SearchDayText}日`);
      await RPA.Logger.info('1日ではないためスルーします');
      ThroughDay.push(SearchDayText);
    } else if (SearchDayText == '1') {
      await RPA.Logger.info(`${SearchDayText}日`);
      await RPA.Logger.info('1日です');
      await RPA.Logger.info('判定処理を開始します');
      break;
    }
  }
  // 1日から判定開始
  const JudgeDayFlag = [];
  JudgeDayFlag[0] = false;
  for (let i = ThroughDay.length + 1; i <= 7; i++) {
    const SearchDay = await RPA.WebBrowser.findElementByXPath(
      `//*[@id="feather-dropdown-7"]/ul/li/div/div[2]/div[2]/div[1]/div[${Number(
        i
      )}]/span`
    );
    const SearchDayText = await SearchDay.getText();
    // 日付にマウスオーバー
    const Day_MouseOver = await RPA.WebBrowser.findElementByXPath(
      `//*[@id="feather-dropdown-7"]/ul/li/div/div[2]/div[2]/div[1]/div[${Number(
        i
      )}]`
    );
    await RPA.WebBrowser.mouseMove(Day_MouseOver);
    await RPA.sleep(50);
    if (Number(SearchDayText) != Number(FilePath_Day)) {
      await RPA.Logger.info('日付が不一致です');
    } else if (Number(SearchDayText) == Number(FilePath_Day)) {
      JudgeDayFlag[0] = true;
      await RPA.Logger.info('日付が一致しました');
      // 一致した日付をクリック
      await RPA.WebBrowser.mouseClick(Day_MouseOver);
      break;
    }
  }
  if (JudgeDayFlag[0] == false) {
    await RPA.Logger.info('次の処理に進みます');
    await JudgeDay(FilePath_Day);
  }

  // カレンダーの2列目から判定再開
  async function JudgeDay(FilePath_Day) {
    const JudgeDayFlag = [];
    JudgeDayFlag[0] = false;
    for (let i = 2; i <= 6; i++) {
      for (let n = 1; n <= 7; n++) {
        const SearchDay = await RPA.WebBrowser.findElementByXPath(
          `//*[@id="feather-dropdown-7"]/ul/li/div/div[2]/div[2]/div[${Number(
            i
          )}]/div[${Number(n)}]/span`
        );
        const SearchDayText = await SearchDay.getText();
        // 日付にマウスオーバー
        const Day_MouseOver = await RPA.WebBrowser.findElementByXPath(
          `//*[@id="feather-dropdown-7"]/ul/li/div/div[2]/div[2]/div[${Number(
            i
          )}]/div[${Number(n)}]`
        );
        await RPA.WebBrowser.mouseMove(Day_MouseOver);
        await RPA.sleep(50);
        if (Number(SearchDayText) == Number(FilePath_Day)) {
          JudgeDayFlag[0] = true;
          await RPA.Logger.info('日付が一致しました');
          // 一致した日付をクリック
          await RPA.WebBrowser.mouseClick(Day_MouseOver);
          break;
        }
      }
      if (JudgeDayFlag[0] == true) {
        await RPA.Logger.info('判定処理を終了します');
        break;
      }
    }
  }

  // 予約時間を入力
  const ReservationTime = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="feather-dropdown-7"]/ul/li/div/div[3]/input'
  );
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await RPA.WebBrowser.sendKeys(ReservationTime, ['8:00am']);
  await ReservationTime.sendKeys(Key.ENTER);
  // 【予約設定する】をクリック
  const ReservationSettings2 = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="root"]/div/div[1]/div/div/div/div/div/div[2]/div/div/div[4]/div/div[2]/button[1]'
  );
  await RPA.WebBrowser.mouseClick(ReservationSettings2);
  await RPA.sleep(5000);
  // 現在のウィンドウを閉じる
  await RPA.WebBrowser.driver.close();
  await RPA.WebBrowser.switchToWindow(windows[1]);
}

async function BrowserReload() {
  const Loading: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('src-components-ErrorMessage-styles-module--loadingErrorMessage')[0].children[0]`
  );
  const LoadingText = await Loading.getText();
  // if (LoadingText == 'このコンテンツの読み込み中に問題が発生しました。') {
  if (LoadingText == 'Sorry, there was a problem loading this content.') {
    await RPA.Logger.info('読み込みエラーが発生したため、ブラウザを更新します');
    await RPA.WebBrowser.refresh();
    await RPA.sleep(5000);
  }
}

async function DeleteFile() {
  // const FirstData = fs.readdirSync(myRPAFolder);
  // await RPA.Logger.info(FirstData);
  // await RPA.Logger.info('予約が完了しましたので、ファイルを削除します');
  // for (let i in FirstData) {
  //   if (FirstData[i].indexOf('.png') > 0) {
  //     fs.unlink(`${myRPAFolder}/${FirstData[i]}`, function(err) {});
  //     break;
  //   }
  // }
}
