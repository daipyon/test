import RPA from 'ts-rpa';
import { WebElement } from 'selenium-webdriver';
const moment = require('moment');

const Kyou = moment().format('M/D');
RPA.Logger.info(Kyou);
const Today = moment().format('YYYY/MM/DD');
RPA.Logger.info(`本日　　　 → ${Today}`);
// 今月の１日をフォーマット変更して取得
const ThisMonthTsuitachi = moment().format('YYYY/MM/1');
RPA.Logger.info(`今月の１日 → ${ThisMonthTsuitachi}`);

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [
  `【Youtube 集計】今月シート（タイトル） の集計が完了しました`
];

// スプレッドシートID
// 本番用
// const SSID = process.env.Senden_Youtube_SheetID3;

// テスト用
const mySSID = process.env.My_SheetID2;

// 今月のシートIDを取得
let ThisMonthSheetID;
// シート名
const SSName = [`News`, `公式`, `バラエティ`, `恋リア`];
// 現在のシート名
let CurrentSSName;

// 作業するスプレッドシートから読み込む行数を記載
let SheetData;
const StartRow = 1;
const LastRow = 200;

// 各番組名の開始行
let ProgramStartRow;
// タイトルの開始行
let TitleStartRow;

// 番組名のリスト
let ProgramList = [];
// 取得した番組名
let CurrentProgramName;
// 取得したタイトル・アップロード日を格納する
let DataList = [];

// エラー発生時のテキストを格納
const ErrorText = [];

async function Start() {
  if (ErrorText.length == 0) {
    // デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
    RPA.Logger.level = 'INFO';
    // 実行前にダウンロードフォルダを全て削除する
    await RPA.File.rimraf({ dirPath: `${process.env.WORKSPACE_DIR}` });
    await RPA.Google.authorize({
      // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    // シートからスプレッドシートIDを取得
    const ThisMonthSheet = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${mySSID}`,
      // 本番用
      // range: `シート1!B13:B13`

      // テスト用
      range: `RPAトリガーシート!B13:B13`
    });
    var id = await ThisMonthSheet[0][0].split(/\//g);
    ThisMonthSheetID = id[5];
    for (let a in SSName) {
      CurrentSSName = SSName[a];
      await RPA.Logger.info(`シート名：【${CurrentSSName}】です`);
      SheetData = await RPA.Google.Spreadsheet.getValues({
        spreadsheetId: ThisMonthSheetID,
        range: `${SSName[a]}!A${StartRow}:CB${LastRow}`
      });
      // B列の番組名をリストに格納
      for (let b in SheetData) {
        await ProgramList.push(SheetData[b][1]);
        if (SheetData[b][1] == `対応者`) {
          //「対応者」という文字を削除してストップ
          await ProgramList.pop();
          break;
        }
      }
      await RPA.Logger.info(ProgramList);
      // B列の最初の3行、空白、undefined などを削除
      ProgramList = await ProgramList.slice(3).filter(v => v);
      await RPA.Logger.info(ProgramList);
      // 番組名の数の分だけ回す
      for (let c in ProgramList) {
        // シートのデータ数の分だけ回す
        for (let d in SheetData) {
          if (ProgramList[c] == SheetData[d][11]) {
            await RPA.Logger.info(`番組名：【${ProgramList[c]}】一致しました`);
            CurrentProgramName = ProgramList[c];
            ProgramStartRow = Number(d) + 3;
            const Title = await RPA.Google.Spreadsheet.getValues({
              spreadsheetId: ThisMonthSheetID,
              range: `${SSName[a]}!N${ProgramStartRow}:CB${LastRow}`
            });
            // N列の本数:1 まで回す
            for (let e in Title) {
              TitleStartRow = ProgramStartRow + Number(e);
              if (Title[e][0] == `1`) {
                await RPA.Logger.info(
                  `${TitleStartRow} 行目　本数:${Title[e][0]} です.ストップします`
                );
                await Work();
                await RPA.Logger.info(
                  `タイトル・視聴回数リストをクリアしました`
                );
                // 格納したタイトルリストを空にする
                DataList = [];
                break;
              }
            }
          }
        }
      }
      // 格納した番組名リストを空にする
      ProgramList = [];
      await RPA.Logger.info(`番組名リストをクリアしました`);
    }
    await RPA.Logger.info(Slack_Text[0]);
    // エラー発生時の処理
    if (ErrorText.length >= 1) {
      // const DOM = await RPA.WebBrowser.driver.getPageSource();
      // await RPA.Logger.info(DOM);
      await RPA.SystemLogger.error(ErrorText);
      Slack_Text[0] = `【Youtube 集計】今月シート（タイトル） でエラーが発生しました！\n${ErrorText}`;
      await RPA.WebBrowser.takeScreenshot();
    }
    await SlackPost(Slack_Text[0]);
    await RPA.WebBrowser.quit();
    await RPA.sleep(1000);
    await process.exit();
  }
}

Start();

let FirstLoginFlag = `true`;
let NotitleFlag = `false`;
async function Work() {
  try {
    // Youtube Studioにログイン
    if (FirstLoginFlag == `true`) {
      await YoutubeLogin();
    }
    if (FirstLoginFlag == `false`) {
      await YoutubeLogin2();
    }
    // 一度ログインしたら、以降はログインページをスキップ
    FirstLoginFlag = `false`;
    // 番組名が【スポット】【今日好き】【恋リアスポット】の場合
    if (
      CurrentProgramName == `スポット`
      // || CurrentProgramName == `今日好き` ||
      // CurrentProgramName == `恋リアスポット`
    ) {
      await RPA.Logger.info(
        `番組名：【${CurrentProgramName}】のためスキップします`
      );
    } else {
      // 番組名を入力
      await SetProgram();
      // 一致する動画がない場合はスキップ
      await NoTitle();
      if (NotitleFlag == `false`) {
        // タイトル・アップロード日を取得
        await GetData();
        // 取得したタイトル・アップロード日をシートに記載
        await SetData();
      }
      // フラグを元に戻す
      NotitleFlag = `false`;
    }
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

// テストは海外版、本番は日本版であることに注意！
async function YoutubeLogin() {
  await RPA.WebBrowser.get(process.env.Youtube_Studio_Url);
  await RPA.sleep(2000);

  // ヘッドレスモードオフ（テスト）用
  // const LoginID = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({ id: 'identifierId' }),
  //   15000
  // );
  // await RPA.WebBrowser.sendKeys(LoginID, [`${process.env.Youtube_Login_ID}`]);
  // const NextButton1 = await RPA.WebBrowser.findElementById('identifierNext');
  // await RPA.WebBrowser.mouseClick(NextButton1);
  // await RPA.sleep(2000);
  // const LoginPW = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({ name: 'password' }),
  //   15000
  // );
  // await RPA.WebBrowser.sendKeys(LoginPW, [`${process.env.Youtube_Login_PW}`]);
  // const NextButton2 = await RPA.WebBrowser.findElementById('passwordNext');
  // await RPA.WebBrowser.mouseClick(NextButton2);

  // 本番・ヘッドレスモードオン（テスト）用
  const LoginID = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'Email' }),
    15000
  );
  await RPA.WebBrowser.sendKeys(LoginID, [`${process.env.Youtube_Login_ID}`]);
  const NextButton1 = await RPA.WebBrowser.findElementById('next');
  await RPA.WebBrowser.mouseClick(NextButton1);
  await RPA.sleep(5000);
  const GoogleLoginPW_1 = await RPA.WebBrowser.findElementsById('Passwd');
  const GoogleLoginPW_2 = await RPA.WebBrowser.findElementsById('password');
  await RPA.Logger.info(`GoogleLoginPW_1 ` + GoogleLoginPW_1.length);
  await RPA.Logger.info(`GoogleLoginPW_2 ` + GoogleLoginPW_2.length);
  if (GoogleLoginPW_1.length == 1) {
    await RPA.WebBrowser.sendKeys(GoogleLoginPW_1[0], [
      `${process.env.Youtube_Login_PW}`
    ]);
    const NextButton2 = await RPA.WebBrowser.findElementById('signIn');
    await RPA.WebBrowser.mouseClick(NextButton2);
  }
  if (GoogleLoginPW_2.length == 1) {
    const GoogleLoginPW = await RPA.WebBrowser.findElementById('password');
    await RPA.WebBrowser.sendKeys(GoogleLoginPW, [
      `${process.env.Youtube_Login_PW}`
    ]);
    const NextButton2 = await RPA.WebBrowser.findElementById('submit');
    await RPA.WebBrowser.mouseClick(NextButton2);
  }

  while (0 == 0) {
    await RPA.sleep(5000);
    const Filter = await RPA.WebBrowser.findElementsByClassName(
      'style-scope ytcp-table-footer'
    );
    if (Filter.length >= 1) {
      await RPA.Logger.info('＊＊＊ログイン成功しました＊＊＊');
      break;
    }
  }
}

async function YoutubeLogin2() {
  while (0 == 0) {
    await RPA.sleep(5000);
    const Filter = await RPA.WebBrowser.findElementsByClassName(
      'style-scope ytcp-table-footer'
    );
    if (Filter.length >= 1) {
      await RPA.Logger.info('＊＊＊ログインをスキップしました＊＊＊');
      break;
    }
  }
}

let SkipFlag = `false`;
let Icon;
let UploadVideo;
async function SetProgram() {
  await RPA.Logger.info(`現在の番組名：【${CurrentProgramName}】`);
  await Wait();
  // アイコンをクリック
  Icon = await RPA.WebBrowser.findElementById('text-input');
  await RPA.WebBrowser.mouseClick(Icon);
  await RPA.sleep(2000);
  if (SkipFlag == `false`) {
    //「動画の日付」をクリック
    // 本番用
    const Videodate = await RPA.WebBrowser.findElementById('text-item-4');

    // テスト用
    // const Videodate = await RPA.WebBrowser.findElementById('text-item-11');
    await RPA.WebBrowser.mouseClick(Videodate);
    await RPA.sleep(1000);
    // そのまま入力するとエラーが起きるため、日付を分割して入力する
    var split = await ThisMonthTsuitachi.split(/\//);
    // 今月の１日を入力
    const StartDate: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByTagName('input')[2]`
    );
    await RPA.WebBrowser.sendKeys(StartDate, [split[0] + '/' + split[1] + '/']);
    await RPA.sleep(500);
    await RPA.WebBrowser.sendKeys(StartDate, [RPA.WebBrowser.Key.BACK_SPACE]);
    await RPA.WebBrowser.sendKeys(StartDate, [RPA.WebBrowser.Key.BACK_SPACE]);
    await RPA.sleep(500);
    await RPA.WebBrowser.sendKeys(StartDate, [split[2]]);
    await RPA.sleep(500);
    // 本日の日付を入力
    var split = await Today.split(/\//);
    const EndDate: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByTagName('input')[3]`
    );
    await RPA.WebBrowser.sendKeys(EndDate, [split[0] + '/' + split[1] + '/']);
    await RPA.sleep(500);
    await RPA.WebBrowser.sendKeys(EndDate, [RPA.WebBrowser.Key.BACK_SPACE]);
    await RPA.WebBrowser.sendKeys(EndDate, [RPA.WebBrowser.Key.BACK_SPACE]);
    await RPA.sleep(500);
    await RPA.WebBrowser.sendKeys(EndDate, [split[2]]);
    // スペースを入れると「適用」ボタンを押すことができる
    await RPA.WebBrowser.sendKeys(EndDate, [RPA.WebBrowser.Key.SPACE]);
    await RPA.sleep(1000);
    //「適用」をクリック
    const Application = await RPA.WebBrowser.findElementById('apply-button');
    await RPA.WebBrowser.mouseClick(Application);
    await RPA.sleep(3000);
    // カーソルが被って視聴回数が取得できないため、画面上部にカーソルを動かす
    UploadVideo = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        id: 'video-list-uploads-tab'
      }),
      15000
    );
    await RPA.WebBrowser.mouseMove(UploadVideo);
    await RPA.sleep(1000);
  }
  // 以降は「動画の日付」の入力をスキップする
  SkipFlag = `true`;
  // 格納した番組名を入力
  await RPA.WebBrowser.sendKeys(Icon, [CurrentProgramName]);
  await RPA.sleep(1000);
  await RPA.WebBrowser.sendKeys(Icon, [RPA.WebBrowser.Key.ENTER]);
  await RPA.sleep(3000);
  // カーソルが被って視聴回数が取得できないため、画面上部にカーソルを動かす
  await RPA.WebBrowser.mouseMove(UploadVideo);
  await RPA.sleep(1000);
}

// 日付検索でヒットした合計数
let Total;
// 1ページ目でデータを取得できたか判定するフラグ
let FisrtGetDataFlag = 'true';
// ループを抜けるためのフラグ
let LoopFlag = 'true';
async function GetData() {
  await Wait();
  // ページ遷移できているか確認のため、現在のページ範囲を取得
  const Range = await RPA.WebBrowser.findElementByClassName(
    'page-description style-scope ytcp-table-footer'
  );
  const RangeText = await Range.getText();
  var text = await RangeText.split(/[\～\/\s]+/);
  Total = Number(text[3]);
  await RPA.Logger.info(`現在のページ範囲　  → ${RangeText}`);
  while (LoopFlag == 'true') {
    // 2ページ目以降は最初にヒットした件数から常に 30 を引いた数だけ回す
    if (FisrtGetDataFlag == 'false') {
      Total = Number(text[3]) - 30;
      await RPA.Logger.info(`残数　　　　　　 → ${Total}`);
    }
    // 検索にヒットした件数分だけ回す
    for (let i = 0; i <= Total - 1; i++) {
      // 日付を取得
      const UpdateDate: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
      );
      const UpdateDateText = await UpdateDate.getText();
      var split = await UpdateDateText.split('\n');
      if (split[1] == '公開日') {
        await RPA.Logger.info(`${split[1]} です.取得します`);
        // ツールのタイトルを取得
        const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${i}].children[0].children[0].children[0].getAttribute('aria-label')`
        );
        await RPA.Logger.info(`タイトル：【${Title}】`);
        await RPA.Logger.info(`公開日　：【${split[0]}】`);
        // 一度リストに格納
        await DataList.push([Title, split[0]]);
      }
      // 1ページ目で取得が完了しなければ次のページに進む
      if (i == 29) {
        FisrtGetDataFlag = 'false';
        await RPA.Logger.info('取得が完了していないため、次のページに進みます');
        const NextPage = await RPA.WebBrowser.findElementById('navigate-after');
        await NextPage.click();
        await RPA.sleep(3000);
        break;
      }
    }
    // 検索にヒットした件数が残り 30 以下になったらループを抜ける
    if (Total < 31) {
      LoopFlag = 'false';
      await RPA.Logger.info(`取得した合計数　 → ${DataList.length}`);
    }
  }
  // フラグを元に戻す
  FisrtGetDataFlag = 'true';
  LoopFlag = 'true';
  // 取得が完了したら×ボタンをクリック
  const DeliteIcon = await RPA.WebBrowser.findElementsByClassName(
    'icon delete-icon style-scope ytcp-chip'
  );
  await RPA.WebBrowser.mouseClick(DeliteIcon[1]);
  await RPA.sleep(1000);
  // 逆順に記載するためリストを反転させる
  await DataList.reverse();
}

async function SetData() {
  for (let i = 0; i <= DataList.length - 1; i++) {
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: ThisMonthSheetID,
      range: `${CurrentSSName}!Q${TitleStartRow - i}:R${TitleStartRow - i}`,
      values: [[DataList[i][0], DataList[i][1]]],
      // 数字の前の ' が入力されなくなる
      parseValues: true
    });
  }
}

// 一覧が出るまで待機する関数
let LoadingFlag = `false`;
async function Wait() {
  // 動画の一覧が表示されるまで回す
  while (LoadingFlag == `false`) {
    const Element = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        id: 'video-title'
      }),
      15000
    );
    const ElementText = await Element.getText();
    if (ElementText.length < 1) {
      await RPA.Logger.info(
        '動画の一覧が表示されないため、ブラウザを更新します'
      );
      await RPA.WebBrowser.refresh();
      await RPA.sleep(1000);
    } else {
      LoadingFlag = `true`;
      await RPA.Logger.info('＊＊＊動画の一覧が表示されました＊＊＊');
    }
  }
  LoadingFlag = `false`;
}

// 一致する動画がない場合はスキップする関数
async function NoTitle() {
  try {
    const NoTitle = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className:
          'no-content-header no-content-header-with-filter style-scope ytcp-video-section-content'
      }),
      15000
    );
    const NoTitleText = await NoTitle.getText();
    if (NoTitleText == '一致する動画がありません。もう一度お試しください。') {
      NotitleFlag = `true`;
      await RPA.Logger.info('一致する動画がないため、取得をスキップします');
      const DeliteIcon = await RPA.WebBrowser.findElementsByClassName(
        'icon delete-icon style-scope ytcp-chip'
      );
      await RPA.WebBrowser.mouseClick(DeliteIcon[1]);
      await RPA.sleep(1000);
    }
  } catch {}
}
