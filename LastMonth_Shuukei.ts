import RPA from 'ts-rpa';
import { WebElement } from 'selenium-webdriver';
const moment = require('moment');

const Today = moment().format('M/D');
RPA.Logger.info(Today);
// 前月の１日をフォーマット変更して取得
let LastMonthTsuitachi = moment()
  .subtract(1, 'months')
  .format('YYYY/MM/1');
RPA.Logger.info(`前月の１日 → ${LastMonthTsuitachi}`);
// 前月末をフォーマット変更して取得
let LastMonthGetsumatsu = moment()
  .subtract(1, 'months')
  .endOf('month')
  .format('YYYY/MM/D');
RPA.Logger.info(`前月の月末 → ${LastMonthGetsumatsu}`);

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Youtube 集計】前月シート の集計が完了しました`];

// RPAトリガーシートのID
const mySSID = process.env.My_SheetID2;
// const SSID = process.env.Senden_Youtube_SheetID3;
// 前月のシートIDを取得
let LastMonthSheetID;
// シート名
const SSName = ['News', '公式', 'バラエティ', 'Mリーグ', '恋リア'];
// const SSName = [`News`];
// 現在のシート名
let CurrentSSName;

// 作業するスプレッドシートから読み込む行数を記載
let SheetData;
const StartRow = 1;
const LastRow = 200;

// 作業対象行と列を取得
let Row;
let Row2;
let Row3 = [];
let StartRow2;
let Column;
let Column2 = 0;

// 番組名のリスト
let ProgramList = [];
// 取得した番組名
let CurrentProgramName;
// タイトルのリスト
let TitleList = [];

// エラー発生時のテキストを格納
const ErrorText = [];

let Count = 0;
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
    const LastMonthSheet = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${mySSID}`,
      // 本番用
      // range: `シート1!B10:B10`

      // テスト用
      range: `RPAトリガーシート!B10:B10`
    });
    var id = await LastMonthSheet[0][0].split(/\//g);
    LastMonthSheetID = id[5];
    // 番組名をリストに格納
    for (let a in SSName) {
      CurrentSSName = SSName[a];
      await RPA.Logger.info(`シート名：【${CurrentSSName}】です`);
      SheetData = await RPA.Google.Spreadsheet.getValues({
        spreadsheetId: LastMonthSheetID,
        range: `${SSName[a]}!A${StartRow}:CB${LastRow}`
      });
      await GetColumn();
      for (let b in SheetData) {
        await ProgramList.push(SheetData[b][1]);
        if (SheetData[b][1] == `` || SheetData[b][1] == undefined) {
          break;
        }
      }
      ProgramList = await ProgramList.slice(3);
      await ProgramList.pop();
      await RPA.Logger.info(ProgramList);
      for (let c in ProgramList) {
        for (let d in SheetData) {
          if (ProgramList[c] == SheetData[d][11]) {
            await RPA.Logger.info(`番組名：【${ProgramList[c]}】一致しました`);
            CurrentProgramName = ProgramList[c];
            Row = Number(d) + 3;
            await RPA.Logger.info(`空白の行 → ${Column} 列 :${Row} 行目`);
            const Title = await RPA.Google.Spreadsheet.getValues({
              spreadsheetId: LastMonthSheetID,
              range: `${SSName[a]}!Q${Row}:CB${LastRow}`
            });
            for (let e in Title) {
              if (Title[e][0] != undefined) {
                Row2 = Row + Number(e);
                await Row3.push(Row2);
                await RPA.Logger.info(`${Column} 列 :${Row2} 行目`);
                if (SheetData[Row2 - 1][Column2] != undefined) {
                  await RPA.Logger.info(
                    `${
                      SheetData[Row2 - 1][Column2]
                    } ← 既に視聴回数が記載されているためスキップします`
                  );
                }
                if (SheetData[Row2 - 1][Column2] == undefined) {
                  {
                    await RPA.Logger.info(Title[e][0]);
                    await TitleList.push(Title[e]);
                  }
                  if (Title[e][0] == ``) {
                    await RPA.Logger.info(
                      `${Row2} 行目 空白です.ストップします`
                    );
                    break;
                  }
                }
              }
            }
          }
        }
        // 最後に「タイトル」や「ここまで」という文字が入るので削除する
        await TitleList.pop();
        await RPA.Logger.info(
          `番組名：【${CurrentProgramName}】タイトル数 → ${TitleList.length}`
        );
        if (TitleList.length == 0) {
          await RPA.Logger.info(
            `番組名：【${CurrentProgramName}】タイトルがないためスキップします`
          );
        }
        if (TitleList.length > 0) {
          await Work();
          await RPA.Logger.info(`タイトル・視聴回数リストをクリアしました`);
        }
        // 格納したタイトルリストを空にする
        TitleList = [];
        Row3 = [];
        Count = 0;
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
      Slack_Text[0] = `【Youtube 集計】でエラーが発生しました！\n${ErrorText}`;
      // await RPA.WebBrowser.takeScreenshot();
    }
    // await SlackPost(Slack_Text[0]);
    await RPA.WebBrowser.quit();
    await RPA.sleep(1000);
    await process.exit();
  }
}

Start();

let FirstLoginFlag = `true`;
let ShityouKaisuu;
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
    if (
      CurrentProgramName == `スポット` ||
      CurrentProgramName == `今日好き` ||
      CurrentProgramName == `恋リアスポット`
    ) {
      await Spot();
      await SetData(ShityouKaisuu);
    } else {
      // 番組名を入力
      await SetProgram();
      // 番組名からタイトルを検索
      await GetData();
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
    var split = await LastMonthTsuitachi.split(/\//);
    // 前月の１日を入力
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
    // 前月末を入力
    var split = await LastMonthGetsumatsu.split(/\//);
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
    // カーソルが被るとテキストが取得できないため、画面上部にカーソルを動かす
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
  // カーソルが被るとテキストが取得できないため、画面上部にカーソルを動かす
  await RPA.WebBrowser.mouseMove(UploadVideo);
  await RPA.sleep(1000);
}

// 番組名が【スポット】もしくは【今日好き】の場合
async function Spot() {
  await RPA.Logger.info(`タイトルごとに視聴回数を取得します`);
  await RPA.Logger.info(`タイトルリスト　　　→ ${TitleList.length} 個`);
  for (let i in TitleList) {
    await Wait();
    // アイコンをクリック
    Icon = await RPA.WebBrowser.findElementById('text-input');
    await RPA.WebBrowser.mouseClick(Icon);
    await RPA.sleep(2000);
    // 格納したタイトルを入力
    await RPA.Logger.info(`シートのタイトル　 → ${TitleList[i][0]}`);
    await RPA.WebBrowser.sendKeys(Icon, [TitleList[i][0]]);
    await RPA.sleep(1000);
    await RPA.WebBrowser.sendKeys(Icon, [RPA.WebBrowser.Key.ENTER]);
    await RPA.sleep(3000);
    // カーソルが被るとテキストが取得できないため、画面上部にカーソルを動かす
    await RPA.WebBrowser.mouseMove(UploadVideo);
    await RPA.sleep(1000);
    // 視聴回数を取得
    const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
    );
    await RPA.Logger.info(`視聴回数　　        → ${NumberofViews}`);
    ShityouKaisuu = NumberofViews;
    StartRow2 = Row3[0] + Count;
    await SetData(ShityouKaisuu);
    Count++;
  }
  await RPA.Logger.info('視聴回数の取得を終了します');
  const DeliteIcon = await RPA.WebBrowser.findElementsByClassName(
    'icon delete-icon style-scope ytcp-chip'
  );
  await RPA.WebBrowser.mouseClick(DeliteIcon[1]);
  await RPA.sleep(1000);
}

let Total;
async function GetData() {
  try {
    // 一致する動画がない場合はスキップ
    await JudgeMatchTitle();
  } catch {
    if (CurrentProgramName == `スポット` || CurrentProgramName == `今日好き`) {
      await RPA.Logger.info(
        `番組名が【${CurrentProgramName}】ですのでスキップします`
      );
    } else {
      await RPA.Logger.info(`タイトルリスト　　　→ ${TitleList.length} 個`);
      await Wait();
      const Range = await RPA.WebBrowser.findElementByClassName(
        'page-description style-scope ytcp-table-footer'
      );
      const RangeText = await Range.getText();
      var text = await RangeText.split(/[\～\/\s]+/);
      Total = Number(text[3]);
      // きちんと遷移できているか確認
      await RPA.Logger.info(`現在のページ範囲　  → ${RangeText}`);
      // 検索にヒットした動画が30件以下の場合
      if (Total < 30) {
        await RPA.Logger.info(`ヒットした件数：${Total} 件 → 30件未満です`);
        await GetData2();
      } else {
        for (let i in TitleList) {
          await RPA.Logger.info(`シートのタイトル　 → ${TitleList[i][0]}`);
          for (let n = 0; n <= 29; n++) {
            // タイトルを取得
            const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${n}].children[0].children[0].children[0].getAttribute('aria-label')`
            );
            await RPA.Logger.info(`ツールのタイトル　 → ${Title}`);
            if (Title == TitleList[i][0]) {
              StartRow2 = Row3[0] + Count;
              await RPA.Logger.info(`タイトル：【${Title}】一致しました`);
              // 視聴回数を取得
              const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${n}].innerText`
              );
              await RPA.Logger.info(`視聴回数　　        → ${NumberofViews}`);
              ShityouKaisuu = NumberofViews;
              await SetData(ShityouKaisuu);
              if (Count == TitleList.length) {
                await RPA.Logger.info(`一致したタイトルの数 → ${Count}`);
                break;
              }
              if (n == 29) {
                await RPA.Logger.info(
                  '視聴回数の取得が完了していないため次のページに進みます'
                );
                const NextPage = await RPA.WebBrowser.findElementById(
                  'navigate-after'
                );
                await NextPage.click();
                await Wait();
              }
              break;
            }
          }
          Count++;
        }
      }
    }
  }
  await RPA.Logger.info('視聴回数の取得を終了します');
  const DeliteIcon = await RPA.WebBrowser.findElementsByClassName(
    'icon delete-icon style-scope ytcp-chip'
  );
  await RPA.WebBrowser.mouseClick(DeliteIcon[1]);
  await RPA.sleep(1000);
}

async function GetData2() {
  for (let i in TitleList) {
    await RPA.Logger.info(`シートのタイトル　 → ${TitleList[i][0]}`);
    for (let n = 0; n <= Total - 1; n++) {
      // タイトルを取得
      const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${n}].children[0].children[0].children[0].getAttribute('aria-label')`
      );
      await RPA.Logger.info(`ツールのタイトル　 → ${Title}`);
      if (Title == TitleList[i][0]) {
        StartRow2 = Row3[0] + Count;
        await RPA.Logger.info(`タイトル：【${Title}】一致しました`);
        // 視聴回数を取得
        const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${n}].innerText`
        );
        await RPA.Logger.info(`視聴回数　　        → ${NumberofViews}`);
        ShityouKaisuu = NumberofViews;
        await SetData(ShityouKaisuu);
        break;
      }
    }
    Count++;
  }
}

// 一覧が出るまで待機する関数
let LoadingFlag = `false`;
async function Wait() {
  while (LoadingFlag == `false`) {
    const Element = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        id: 'video-title'
      }),
      15000
    );
    const ElementText = await Element.getText();
    if (ElementText.length < 1) {
      await RPA.Logger.info('動画の一覧が出ないため、ブラウザを更新します');
      await RPA.WebBrowser.refresh();
      await RPA.sleep(1000);
    } else {
      LoadingFlag = `true`;
      await RPA.Logger.info('＊＊＊動画の一覧が出ました＊＊＊');
    }
  }
  LoadingFlag = `false`;
}

async function SetData(ShityouKaisuu) {
  await RPA.Logger.info(`${Column} 列 :${StartRow2} 行目に記載します`);
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: LastMonthSheetID,
    range: `${CurrentSSName}!${Column}${StartRow2}:${Column}${StartRow2}`,
    values: [[ShityouKaisuu]],
    // 数字の前の ' が入力されなくなる
    parseValues: true
  });
}

async function JudgeMatchTitle() {
  const NoTitle = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className:
        'no-content-header no-content-header-with-filter style-scope ytcp-video-section-content'
    }),
    15000
  );
  const NoTitleText = await NoTitle.getText();
  if (NoTitleText == '一致する動画がありません。もう一度お試しください。') {
    await RPA.Logger.info('一致する動画がないため、取得をスキップします');
  }
}

const ColumnList = [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I',
  'J',
  'K',
  'L',
  'M',
  'N',
  'O',
  'P',
  'Q',
  'R',
  'S',
  'T',
  'U',
  'V',
  'W',
  'X',
  'Y',
  'Z',
  'AA',
  'AB',
  'AC',
  'AD',
  'AE',
  'AF',
  'AG',
  'AH',
  'AI',
  'AJ',
  'AK',
  'AL',
  'AM',
  'AN',
  'AO',
  'AP',
  'AQ',
  'AR',
  'AS',
  'AT',
  'AU',
  'AV',
  'AW',
  'AX',
  'AY',
  'AZ',
  'BA',
  'BB',
  'BC',
  'BD',
  'BE',
  'BF',
  'BG',
  'BH',
  'BI',
  'BJ',
  'BK',
  'BL',
  'BM',
  'BN',
  'BO',
  'BP',
  'BQ',
  'BR',
  'BS',
  'BT',
  'BU',
  'BV',
  'BW',
  'BX',
  'BY',
  'BZ',
  'CA',
  'CB'
];
async function GetColumn() {
  // 記載する開始列を取得
  for (let i in SheetData) {
    if (Today == SheetData[2][i]) {
      Column = ColumnList[i];
      break;
    }
    if (Today != SheetData[2][i]) {
      Column2++;
    }
  }
}
