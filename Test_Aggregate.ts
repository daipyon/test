import RPA from 'ts-rpa';
import { WebElement } from 'selenium-webdriver';
const moment = require('moment');

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Youtube 集計】集計完了しました`];

// RPAトリガーシートのID
const mySSID = process.env.My_SheetID2;
const mySSID2 = [];
// const SSID = process.env.Senden_Youtube_SheetID3;
// const SSID2 = [];
// シート名
// const SSName4 = ['News', '公式', 'バラエティ', 'Mリーグ'];

// 6/1 追加修正
const SSName4 = ['News', '公式', 'バラエティ', 'Mリーグ', '恋リア'];
// const SSName4 = ['News', 'Mリーグ', '恋リア'];

// 作業するスプレッドシートから読み込む行数を記載
const StartRow = 4;
const LastRow = 100;
const StartRow2 = 2;
const LastRow2 = 200;
const StartRow3 = 13;

// 現在のシートIDとシート名
const CurrentSSID = [];
const CurrentSSName = [];

// 作業対象行とデータを取得
let WorkData;
let Row;
let Row2;
let Row3;
let Row4;

// 日付
let Today;
let NewToday;
let Yesterday;
let ThisMonth;
let TwoMonthsBefore;
let BeginningOfMonth;
let date = [];

// 番組名のリスト
let ProgramList = [];
// 取得した番組名を格納
let ProgramName = [];
// タイトルのリスト
let TitleList = [];
let UploadDateList = [];

// エラー発生時のテキストを格納
const ErrorText = [];

// 今月分からスタートする場合はここをコメントアウト
let MonthFlag = '前月';
// 前月分からスタートする場合はここをコメントアウト
// let MonthFlag = '今月';

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
    // シートからスプレッドシートIDを取得
    const LastMonthSheet = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${mySSID}`,
      // テスト用
      range: `RPAトリガーシート!B10:B10`

      // 本番用
      // range: `シート1!B10:B10`
    });
    const ThisMonthSheet = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${mySSID}`,
      // テスト用
      range: `RPAトリガーシート!B13:B13`

      // 本番用
      // range: `シート1!B13:B13`
    });
    // 5/25 IDのみ取得
    var id = await LastMonthSheet[0][0].split(/\//g);
    const LastMonthSheetID = id[5];
    var id2 = await ThisMonthSheet[0][0].split(/\//g);
    const ThisMonthSheetID = id2[5];
    // 今月分からスタートする場合は上をコメントアウト
    await mySSID2.push(LastMonthSheetID);
    await mySSID2.push(ThisMonthSheetID);
    await RPA.Logger.info(mySSID2);
    await RPA.Logger.info(`【${MonthFlag}】分のシートです`);
    if (MonthFlag == '前月') {
      await RPA.Logger.info('視聴回数のみ取得を開始します');
    }
    // テスト用
    if (MonthFlag == '今月') {
      await RPA.Logger.info(
        'タイトル、アップロード日、視聴回数の取得を開始します'
      );
      // 5/25 【恋リア】シート追加 → 6/1 コメントアウト
      // await SSName4.push('恋リア');
      // await RPA.Logger.info('【恋リア】シートを追加しました');
    }
    // 番組名をリストに格納
    for (let j in mySSID2) {
      for (let n in SSName4) {
        CurrentSSID[0] = mySSID2[j];
        CurrentSSName[0] = SSName4[n];
        const Program = await RPA.Google.Spreadsheet.getValues({
          spreadsheetId: `${mySSID2[j]}`,
          range: `${SSName4[n]}!B${StartRow}:B${LastRow}`
        });
        // 6/10 追加修正
        if (MonthFlag == '今月' && CurrentSSName[0] == 'Mリーグ') {
          await RPA.Logger.info(
            `番組名が【${CurrentSSName[0]}】のためスキップします`
          );
        } else {
          // 6/1 追加修正
          if (MonthFlag == '今月' && CurrentSSName[0] == '恋リア') {
            await RPA.Logger.info(`シート名が【${CurrentSSName[0]}】です`);
            await ProgramList.push(`恋リアスポット`);
          } else {
            for (let i in Program) {
              if (Program[i][0] != undefined) {
                await ProgramList.push(Program[i][0]);
              } else {
                break;
              }
            }
          }
          // M列(本日の日付)を取得
          if (ProgramList[0] == 'アベプラ') {
            Today = await RPA.Google.Spreadsheet.getValues({
              spreadsheetId: `${mySSID2[j]}`,
              range: `${SSName4[n]}!M1:M1`
            });
            // 本日の日付をフォーマット変更して取得
            NewToday = moment().format('YYYY/MM/DD');
            await RPA.Logger.info('本日の日付　　　　　→ ' + NewToday);
            // 前日の日付をフォーマット変更して取得
            Yesterday = moment()
              .add(-1, 'days')
              .format('YYYY/MM/DD');
            await RPA.Logger.info('前日の日付　　　　　→ ' + Yesterday);
            // 今月の日付をフォーマット変更して取得
            ThisMonth = moment().format('YYYY/MM');
            await RPA.Logger.info('今月　　　　　　　　→ ' + ThisMonth);
            // 前々月の日付をフォーマット変更して取得
            TwoMonthsBefore = moment()
              .add(-2, 'month')
              .format('YYYY/MM');
            await RPA.Logger.info('前々月　　　　　　　→ ' + TwoMonthsBefore);
            // 6/1 追加修正
            // 月初めの日付をフォーマット変更して取得
            BeginningOfMonth = moment().format('M/1');
            await RPA.Logger.info('月初め　　　　　　　→ ' + BeginningOfMonth);
          }
          date = await RPA.Google.Spreadsheet.getValues({
            spreadsheetId: `${mySSID2[j]}`,
            range: `${SSName4[n]}!Q3:CA${LastRow2}`
          });
          const ProgramName2 = await RPA.Google.Spreadsheet.getValues({
            spreadsheetId: `${mySSID2[j]}`,
            range: `${SSName4[n]}!L${StartRow2}:L${LastRow2}`
          });
          for (let v in ProgramList) {
            for (let k in ProgramName2) {
              if (ProgramName2[k][0] == ProgramList[v]) {
                // 格納したリストを空にする
                TitleList = [];
                // 6/9 追加修正
                UploadDateList = [];
                Row = Number(k) + StartRow;
                const JudgeTitle = await RPA.Google.Spreadsheet.getValues({
                  spreadsheetId: `${mySSID2[j]}`,
                  range: `${SSName4[n]}!Q${Row}:Q${Row}`
                });
                if (JudgeTitle == undefined) {
                  let JudgeTitle2 = await RPA.Google.Spreadsheet.getValues({
                    spreadsheetId: `${mySSID2[j]}`,
                    range: `${SSName4[n]}!Q${Row}:Q${LastRow2}`
                  });
                  for (let i = 0; i <= JudgeTitle2.length - 1; i++) {
                    if (JudgeTitle2[i][0] != undefined) {
                      Row3 = Row + i;
                      // 開始行から空白が１つ以上ある場合
                      await RPA.Logger.info(`${Row3} 行目 タイトルがあります`);
                      await RPA.Logger.info(`${Row3} 行目から取得を開始します`);
                      // タイトルがない場合
                      if (
                        JudgeTitle2[i][0] == 'タイトル' ||
                        JudgeTitle2[i][0] == 'ここまで'
                      ) {
                        await RPA.Logger.info(`パターン 4 です`);
                        await RPA.Logger.info(`タイトルがありませんでした`);
                        break;
                      }
                      break;
                    }
                  }
                  // パターン1の終了行を取得
                  JudgeTitle2 = await RPA.Google.Spreadsheet.getValues({
                    spreadsheetId: `${mySSID2[j]}`,
                    range: `${SSName4[n]}!Q${Row3}:Q${LastRow2}`
                  });
                  for (let i = 0; i <= JudgeTitle2.length - 1; i++) {
                    if (JudgeTitle2[i][0] == undefined) {
                      Row4 = Row3 + i - 1;
                      await RPA.Logger.info(`パターン 1 です`);
                      await RPA.Logger.info(
                        `パターン 1 ${Row4} 行目までがタイトルです`
                      );
                      break;
                    }
                    // 6/8 追加修正
                    if (JudgeTitle2[i][0] == 'ここまで') {
                      Row4 = Row3 + i;
                      await RPA.Logger.info(`パターン 5 です`);
                      await RPA.Logger.info(
                        `パターン 5 ${Row4} 行目が"ここまで"です`
                      );
                      break;
                    }
                  }
                } else {
                  Row2 = Row + 1;
                  await RPA.Logger.info(`次の行が空白かどうかチェックします`);
                  let JudgeTitle3 = await RPA.Google.Spreadsheet.getValues({
                    spreadsheetId: `${mySSID2[j]}`,
                    range: `${SSName4[n]}!Q${Row2}:Q${Row2}`
                  });
                  if (JudgeTitle3 != undefined) {
                    Row3 = Row;
                    // 開始行から空白が１つもない場合
                    await RPA.Logger.info(`パターン 2 です`);
                    await RPA.Logger.info(`${Row3} 行目 タイトルがあります`);
                    await RPA.Logger.info(`${Row3} 行目から取得を開始します`);
                    // パターン2の終了行を取得
                    JudgeTitle3 = await RPA.Google.Spreadsheet.getValues({
                      spreadsheetId: `${mySSID2[j]}`,
                      range: `${SSName4[n]}!Q${Row3}:Q${LastRow2}`
                    });
                    for (let i = 0; i <= JudgeTitle3.length - 1; i++) {
                      if (JudgeTitle3[i][0] == undefined) {
                        Row4 = Row3 + i - 1;
                        await RPA.Logger.info(
                          `パターン 2 ${Row4} 行目までがタイトルです`
                        );
                        break;
                      }
                    }
                  } else {
                    await RPA.Logger.info(`タイトルがある行をチェックします`);
                    let JudgeTitle4 = await RPA.Google.Spreadsheet.getValues({
                      spreadsheetId: `${mySSID2[j]}`,
                      range: `${SSName4[n]}!Q${Row2}:Q${LastRow2}`
                    });
                    for (let i = 0; i <= JudgeTitle4.length - 1; i++) {
                      if (JudgeTitle4[i][0] != undefined) {
                        Row3 = Row2 + i;
                        // 開始行にタイトルがあるが、次のタイトルから空白が１つ以上ある場合（イレギュラー）
                        await RPA.Logger.info(`パターン 3 です`);
                        await RPA.Logger.info(
                          `${Row3} 行目 タイトルがあります`
                        );
                        await RPA.Logger.info(
                          `${Row3} 行目から取得を開始します`
                        );
                        // パターン3の終了行を取得
                        JudgeTitle4 = await RPA.Google.Spreadsheet.getValues({
                          spreadsheetId: `${mySSID2[j]}`,
                          range: `${SSName4[n]}!Q${Row3}:Q${LastRow2}`
                        });
                        for (let i = 0; i <= JudgeTitle4.length - 1; i++) {
                          if (JudgeTitle4[i][0] == undefined) {
                            Row4 = Row3 + i - 1;
                            await RPA.Logger.info(
                              `パターン 3 ${Row4} 行目までがタイトルです`
                            );
                            break;
                          }
                        }
                        break;
                      }
                    }
                  }
                }
                // 6/9 追加修正
                WorkData = await RPA.Google.Spreadsheet.getValues({
                  spreadsheetId: `${mySSID2[j]}`,
                  range: `${SSName4[n]}!Q${Row3}:R${Row4}`
                });
                // タイトルリストに格納
                for (let i in WorkData) {
                  await TitleList.push(WorkData[i][0]);
                  await UploadDateList.push(WorkData[i][1]);
                  if (
                    TitleList[0] == 'タイトル' ||
                    TitleList[0] == 'ここまで'
                  ) {
                    await RPA.Logger.info(
                      'タイトルがないためリスト内を空にします'
                    );
                    TitleList = [];
                    UploadDateList = [];
                    break;
                  }
                }
                if (ProgramName[0] < 1) {
                  await ProgramName.push(ProgramList[v]);
                } else {
                  // 格納した番組名の先頭を削除
                  await ProgramName.shift();
                  await ProgramName.push(ProgramList[v]);
                }
                await Work();
                break;
              }
            }
          }
          // 格納した番組名リストを空にする
          ProgramList = [];
          await RPA.Logger.info('番組名リストをクリアします');
          await RPA.Logger.info(ProgramList);
        }
      }
      if (Number(j) == 0) {
        // 今月分からスタートする場合はここをコメントアウト
        MonthFlag = '今月';
        await RPA.Logger.info(`【${MonthFlag}】分のシートに移動します`);
        await RPA.Logger.info(
          'タイトル、アップロード日、視聴回数の取得を開始します'
        );
        // 5/25 【恋リア】シート追加 → 6/1 コメントアウト
        // await SSName4.push('恋リア');
        // await RPA.Logger.info('【恋リア】シートを追加しました');
        // 前月分からスタートする場合はここをコメントアウト
        // await RPA.Logger.info('作業を終了します');
      }
      if (Number(j) == 1) {
        // 6/10 追加修正
        // 格納したリストを空にする
        TitleList = [];
        UploadDateList = [];
        ProgramName[0] = `Mリーグ`;
        await RPA.Logger.info(`ログアウトしてアカウントを切り替えます`);
        const AccountButton = await RPA.WebBrowser.findElementById(
          `account-button`
        );
        await AccountButton.click();
        await RPA.sleep(1000);
        await RPA.WebBrowser.wait(
          RPA.WebBrowser.Until.elementLocated({
            className: `style-scope ytcp-popup-container`
          }),
          15000
        );
        const Logout: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('yt-simple-endpoint style-scope ytd-compact-link-renderer')[2]`
        );
        await Logout.click();
        await RPA.sleep(5000);
        const Login = await RPA.WebBrowser.wait(
          RPA.WebBrowser.Until.elementLocated({
            className: `style-scope ytd-button-renderer style-suggestive size-small`
          }),
          15000
        );
        await Login.click();
        await RPA.sleep(3000);
        await Work2();
        await RPA.Logger.info('作業を終了します');
      }
    }
  }
  // エラー発生時の処理
  if (ErrorText.length >= 1) {
    // const DOM = await RPA.WebBrowser.driver.getPageSource();
    // await RPA.Logger.info(DOM);
    await RPA.SystemLogger.error(ErrorText);
    Slack_Text[0] = `【Youtube 集計】でエラーが発生しました！\n${ErrorText}`;
    await RPA.WebBrowser.takeScreenshot();
  }
  await SlackPost(Slack_Text[0]);
  await RPA.WebBrowser.quit();
  await RPA.sleep(1000);
  await process.exit();
}

Start();

let FirstLoginFlag = 'true';
let NumberofViewsList = [];
async function Work() {
  try {
    await RPA.Logger.info('データ取得を開始します');
    // Youtube Studioにログイン
    if (FirstLoginFlag == 'true') {
      await YoutubeLogin();
    }
    if (FirstLoginFlag == 'false') {
      await YoutubeLogin2();
    }
    // 一度ログインしたら、以降はログインページをスキップ
    FirstLoginFlag = 'false';
    // 5/25 追加修正 → 6/1 追加修正２
    // if (ProgramName[0] == 'スポット' || ProgramName[0] == '今日好き') {
    if (
      ProgramName[0] == 'スポット' ||
      ProgramName[0] == '今日好き' ||
      ProgramName[0] == '恋リアスポット'
    ) {
      await RPA.Logger.info(`番組名が【${ProgramName[0]}】です`);
      // 5/18 追加修正
      await Spot(TitleList, NumberofViewsList);
      await SetData();
    } else {
      // 番組名を入力
      await SetProgram();
      // 番組名からタイトルを検索
      await GetData();
      // シートに取得したデータを記載
      await SetData();
    }
  } catch (error) {
    ErrorText[0] = error;
    await Start();
  }
}

// 6/10 追加修正
async function Work2() {
  try {
    await RPA.Logger.info(`【${ProgramName[0]}】のデータ取得を開始します`);
    await YoutubeLogin();
    await MLeagueGetData();
    await MLeagueSetData();
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
  await RPA.Logger.info('Youtube Studioにログインします');
  if (MonthFlag == '今月' && ProgramName[0] == 'Mリーグ') {
    // 6/12 追加修正
    // テスト用
    // const OtherAccount: WebElement = await RPA.WebBrowser.driver.executeScript(
    //   `return document.getElementsByClassName('lCoei YZVTmd SmR8')[1]`
    // );

    // 本番用
    const OtherAccount = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({ className: `tXIHtc` }),
      15000
    );
    await OtherAccount.click();
  } else {
    await RPA.WebBrowser.get(process.env.Youtube_Studio_Url);
  }
  await RPA.sleep(2000);
  await RPA.WebBrowser.takeScreenshot();

  // ヘッドレスモードオフ（テスト）用
  // const LoginID = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({ id: 'identifierId' }),
  //   8000
  // );
  // if (MonthFlag == '今月' && ProgramName[0] == 'Mリーグ') {
  //   await RPA.WebBrowser.sendKeys(LoginID, [
  //     `${process.env.Youtube_Login_MLeagueID}`
  //   ]);
  // } else {
  //   await RPA.WebBrowser.sendKeys(LoginID, [`${process.env.Youtube_Login_ID}`]);
  // }
  // const NextButton1 = await RPA.WebBrowser.findElementById('identifierNext');
  // await RPA.WebBrowser.mouseClick(NextButton1);
  // await RPA.sleep(5000);
  // const LoginPW = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({ name: 'password' }),
  //   8000
  // );
  // if (MonthFlag == '今月' && ProgramName[0] == 'Mリーグ') {
  //   await RPA.WebBrowser.sendKeys(LoginPW, [
  //     `${process.env.Youtube_Login_MLeaguePW}`
  //   ]);
  // } else {
  //   await RPA.WebBrowser.sendKeys(LoginPW, [`${process.env.Youtube_Login_PW}`]);
  // }
  // const NextButton2 = await RPA.WebBrowser.findElementById('passwordNext');
  // await RPA.WebBrowser.mouseClick(NextButton2);

  // 本番・ヘッドレスモードオン（テスト）用
  const LoginID = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'Email' }),
    8000
  );
  if (MonthFlag == '今月' && ProgramName[0] == 'Mリーグ') {
    await RPA.WebBrowser.sendKeys(LoginID, [
      `${process.env.Youtube_Login_MLeagueID}`
    ]);
  } else {
    await RPA.WebBrowser.sendKeys(LoginID, [`${process.env.Youtube_Login_ID}`]);
  }
  await RPA.WebBrowser.takeScreenshot();
  const NextButton1 = await RPA.WebBrowser.findElementById('next');
  await RPA.WebBrowser.mouseClick(NextButton1);
  await RPA.sleep(5000);
  await RPA.WebBrowser.takeScreenshot();
  const GoogleLoginPW_1 = await RPA.WebBrowser.findElementsById('Passwd');
  const GoogleLoginPW_2 = await RPA.WebBrowser.findElementsById('password');
  await RPA.Logger.info(`GoogleLoginPW_1 ` + GoogleLoginPW_1.length);
  await RPA.Logger.info(`GoogleLoginPW_2 ` + GoogleLoginPW_2.length);
  // 6/13 上記どちらも0だった
  if (MonthFlag == '今月' && ProgramName[0] == 'Mリーグ') {
    const DOM = await RPA.WebBrowser.driver.getPageSource();
    await RPA.Logger.info(DOM);
    await RPA.WebBrowser.takeScreenshot();
  }
  if (GoogleLoginPW_1.length == 1) {
    if (MonthFlag == '今月' && ProgramName[0] == 'Mリーグ') {
      await RPA.WebBrowser.sendKeys(GoogleLoginPW_1[0], [
        `${process.env.Youtube_Login_MLeaguePW}`
      ]);
    } else {
      await RPA.WebBrowser.sendKeys(GoogleLoginPW_1[0], [
        `${process.env.Youtube_Login_PW}`
      ]);
    }
    const NextButton2 = await RPA.WebBrowser.findElementById('signIn');
    await RPA.WebBrowser.mouseClick(NextButton2);
  }
  if (GoogleLoginPW_2.length == 1) {
    const GoogleLoginPW = await RPA.WebBrowser.findElementById('password');
    if (MonthFlag == '今月' && ProgramName[0] == 'Mリーグ') {
      await RPA.WebBrowser.sendKeys(GoogleLoginPW, [
        `${process.env.Youtube_Login_MLeaguePW}`
      ]);
    } else {
      await RPA.WebBrowser.sendKeys(GoogleLoginPW, [
        `${process.env.Youtube_Login_PW}`
      ]);
    }
    const NextButton2 = await RPA.WebBrowser.findElementById('submit');
    await RPA.WebBrowser.mouseClick(NextButton2);
  }

  // 6/10 追加修正
  if (MonthFlag == '今月' && ProgramName[0] == 'Mリーグ') {
    while (0 == 0) {
      await RPA.sleep(5000);
      const Filter = await RPA.WebBrowser.findElementsByClassName(
        'style-scope ytd-two-column-browse-results-renderer'
      );
      if (Filter.length >= 1) {
        await RPA.Logger.info('＊＊＊ログイン成功しました＊＊＊');
        break;
      }
    }
    await RPA.WebBrowser.get(process.env.Youtube_MLeague_Url);
  }
  while (0 == 0) {
    await RPA.sleep(5000);
    const Filter = await RPA.WebBrowser.findElementsByClassName(
      'style-scope ytcp-table-footer'
    );
    if (Filter.length >= 1) {
      if (MonthFlag == '今月' && ProgramName[0] == 'Mリーグ') {
        await RPA.Logger.info(`＊＊＊${ProgramName[0]} の管理ページです＊＊＊`);
      } else {
        await RPA.Logger.info('＊＊＊ログイン成功しました＊＊＊');
      }
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

async function SetProgram() {
  await RPA.Logger.info(`取得した番組名　    → 【${ProgramName[0]}】`);
  await Wait();
  const Icon = await RPA.WebBrowser.findElementById('text-input');
  await RPA.WebBrowser.mouseClick(Icon);
  await RPA.sleep(2000);

  // 「タイトル」をクリック
  // ヘッドレスモードオン（テスト）用
  // const Title = await RPA.WebBrowser.findElementById('text-item-5');

  // 本番用・ヘッドレスモードオフ（テスト）用
  // const Title = await RPA.WebBrowser.findElementById('text-item-0');
  // 6/10 要素変更
  const Title = await RPA.WebBrowser.findElementById('text-item-1');

  await RPA.WebBrowser.mouseClick(Title);
  await RPA.sleep(1000);

  // 格納した番組名を入力
  const InputTitle: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByTagName('input')[2]`
  );
  await RPA.WebBrowser.sendKeys(InputTitle, [ProgramName[0]]);
  await RPA.sleep(1000);
  // 「適用」をクリック
  const Application = await RPA.WebBrowser.findElementById('apply-button');
  await RPA.WebBrowser.mouseClick(Application);
  await RPA.sleep(3000);
  // カーソルが被るとテキストが取得できないため、画面上部にカーソルを動かす
  const UploadVideo = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      id: 'video-list-uploads-tab'
    }),
    15000
  );
  await RPA.WebBrowser.mouseMove(UploadVideo);
  await RPA.sleep(1000);
}

// 番組名が【スポット】もしくは【今日好き】の場合の処理
async function Spot(TitleList, NumberofViewsList) {
  await RPA.Logger.info('タイトルごとに視聴回数を取得します');
  await RPA.Logger.info('タイトルリスト数　　→', TitleList.length, '個');
  await RPA.Logger.info('タイトルリスト　　　→', TitleList);
  for (let i in TitleList) {
    await Wait();
    // アイコンをクリック
    const Icon = await RPA.WebBrowser.findElementById('text-input');
    await RPA.WebBrowser.mouseClick(Icon);
    await RPA.sleep(2000);

    // 「タイトル」をクリック
    // ヘッドレスモードオン（テスト）用
    // const Title = await RPA.WebBrowser.findElementById('text-item-5');

    // 本番用・ヘッドレスモードオフ（テスト）用
    // const Title = await RPA.WebBrowser.findElementById('text-item-0');
    // 6/10 要素変更
    const Title = await RPA.WebBrowser.findElementById('text-item-1');

    await RPA.WebBrowser.mouseClick(Title);
    await RPA.sleep(1000);

    // 格納した番組名を入力
    const InputTitle: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByTagName('input')[2]`
    );
    await RPA.WebBrowser.sendKeys(InputTitle, [TitleList[i]]);
    await RPA.sleep(1000);
    // 「適用」をクリック
    const Application = await RPA.WebBrowser.findElementById('apply-button');
    await RPA.WebBrowser.mouseClick(Application);
    await RPA.sleep(3000);
    const UploadVideo = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        id: 'video-list-uploads-tab'
      }),
      15000
    );
    await RPA.WebBrowser.mouseMove(UploadVideo);
    await RPA.sleep(1000);
    // 一致する動画がない場合はスキップ
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
        await RPA.Logger.info('一致する動画がないため、取得をスキップします');
      }
    } catch {
      await Wait();
      // 6/9 追加修正
      // const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
      //   `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
      // );
      // const TitleText = await Title.getText();
      // await RPA.Logger.info('ツールのタイトル　　→', TitleText);
      // const date: WebElement = await RPA.WebBrowser.driver.executeScript(
      //   `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
      // );
      // let dateText;
      // dateText = await date.getText();
      // dateText = await dateText.split('\n');
      // await RPA.Logger.info('ツールの日付　　　　→', dateText[0]);
      // // タイトルとアップロード日が同じ視聴回数を取得
      // if (TitleText == TitleList[i] && dateText[0] == UploadDateList[i]) {
      //   const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
      //     `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
      //   );
      //   await RPA.Logger.info('視聴回数　　        →', NumberofViews);
      //   await NumberofViewsList.push(NumberofViews);
      // } else if (TitleText == TitleList[i]) {
      const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[0].innerText`
      );
      await RPA.Logger.info('視聴回数　　        →', NumberofViews);
      await NumberofViewsList.push(NumberofViews);
      // }
    }
    const DeliteIcon = await RPA.WebBrowser.findElementById('delete-icon');
    await RPA.WebBrowser.mouseClick(DeliteIcon);
    await RPA.sleep(3000);
  }
  await RPA.Logger.info(`【${ProgramName[0]}】のデータ取得を終了します`);
  await RPA.Logger.info('視聴回数リスト数　　→', NumberofViewsList.length);
}

let LoopFlag = 'true';
let Total = [];
let YoutubeTitleList = [];
let MatchTitleList = [];
let YesterdayUpdateTitleList = [];
let YesterdayNumberofViewsList = [];
let TodayUpdateTitleList = [];
let TodayNumberofViewsList = [];
// 6/10 追加修正
async function MLeagueGetData() {
  for (let i = 0; i <= 29; i++) {
    const UpdateDate: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
    );
    const UpdateDateText = await UpdateDate.getText();
    var split = await UpdateDateText.split('\n');
    if (
      (split[0].includes(ThisMonth) == true && split[1] == '公開日') ||
      (split[0].includes(ThisMonth) == true && split[1] == '公開予約')
    ) {
      await RPA.Logger.info(
        '今月のタイトル・アップロード日・視聴回数を取得します'
      );
      // タイトルを取得
      const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
      );
      const TitleText = await Title.getText();
      await YoutubeTitleList.push(TitleText);
      // アップロード日を取得
      await UploadDateList.push(split[0]);
      // 視聴回数を取得
      const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
      );
      await NumberofViewsList.push(NumberofViews);
    }
  }
  await RPA.Logger.info('ツールのリスト　　　→', YoutubeTitleList);
  await RPA.Logger.info('アップロード日　　　→', UploadDateList);
}

async function GetData() {
  // 一致する動画がない場合はスキップ
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
      await RPA.Logger.info('一致する動画がないため、取得をスキップします');
    }
  } catch {
    if (ProgramName[0] == 'スポット' || ProgramName[0] == '今日好き') {
      await RPA.Logger.info(
        `番組名が【${ProgramName[0]}】ですのでスキップします`
      );
    } else {
      await RPA.Logger.info('タイトルリスト数　　→', TitleList.length, '個');
      // アップロード日が前々月のタイトルを取得した時点で処理を抜ける
      while (0 == 0) {
        await Wait();
        const Range = await RPA.WebBrowser.findElementByClassName(
          'page-description style-scope ytcp-table-footer'
        );
        const RangeText = await Range.getText();
        var text = await RangeText.split(/[\～\/\s]+/);
        // きちんと遷移できているか確認
        await RPA.Logger.info('現在のページ範囲　  →', RangeText);
        Total[0] = Number(text[3]);
        // 検索にヒットした動画が30件以下の場合
        if (Total[0] < 30) {
          for (let i = 0; i <= Total[0] - 1; i++) {
            // 以下だとなぜかページの上から6番目のタイトルが取得不可
            // const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
            //   `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
            // );
            // const TitleText = await Title.getText();
            // タイトルを取得
            const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${i}].children[0].children[0].children[0].getAttribute('aria-label')`
            );
            await YoutubeTitleList.push(Title);
            // 日付を取得
            const UpdateDate: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
            );
            const UpdateDateText = await UpdateDate.getText();
            var split = await UpdateDateText.split('\n');
            if (split[0] == Yesterday && MonthFlag == '今月') {
              await RPA.Logger.info(
                '前日の公開日が一致したタイトルを取得します'
              );
              const Title2: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
              );
              const TitleText = await Title2.getText();
              // 5/25 日付の下の文字を取得
              const date: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
              );
              const JudgeText = await date.getText();
              var split = await JudgeText.split('\n');
              // 視聴回数を取得（文字列だと数字の前に ' が付きシートの関数が効かないため Number型に変更する）
              const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
              );
              if (split[1] != '公開日') {
                await RPA.Logger.info(split[1]);
                await RPA.Logger.info(
                  '"公開日"ではないため取得をスキップします'
                );
              } else {
                await RPA.Logger.info(split[1], 'です');
                await YesterdayUpdateTitleList.push(TitleText);
                await RPA.Logger.info('視聴回数　　        →', NumberofViews);
                await YesterdayNumberofViewsList.push(NumberofViews);
              }
            }
            if (split[0] == NewToday && MonthFlag == '今月') {
              await RPA.Logger.info(
                '本日の公開日が一致したタイトルを取得します'
              );
              const Title2: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
              );
              const TitleText = await Title2.getText();
              // 5/25 日付の下の文字を取得
              const date: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
              );
              const JudgeText = await date.getText();
              var split = await JudgeText.split('\n');
              // 視聴回数を取得（文字列だと数字の前に ' が付きシートの関数が効かないため Number型に変更する）
              const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
              );
              if (split[1] != '公開日') {
                await RPA.Logger.info(split[1]);
                await RPA.Logger.info(
                  '"公開日"ではないため取得をスキップします'
                );
              } else {
                await RPA.Logger.info(split[1], 'です');
                await TodayUpdateTitleList.push(TitleText);
                await RPA.Logger.info('視聴回数　　        →', NumberofViews);
                await TodayNumberofViewsList.push(NumberofViews);
              }
            }
          }
        } else {
          for (let i = 0; i <= 29; i++) {
            // タイトルを取得
            const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${i}].children[0].children[0].children[0].getAttribute('aria-label')`
            );
            await YoutubeTitleList.push(Title);
            // 日付を取得
            const UpdateDate: WebElement = await RPA.WebBrowser.driver.executeScript(
              `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
            );
            const UpdateDateText = await UpdateDate.getText();
            var split = await UpdateDateText.split('\n');
            if (split[0] == Yesterday && MonthFlag == '今月') {
              await RPA.Logger.info(
                '前日の公開日が一致したタイトルを取得します'
              );
              const Title2: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
              );
              const TitleText = await Title2.getText();
              // 5/25 日付の下の文字を取得
              const date: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
              );
              const JudgeText = await date.getText();
              var split = await JudgeText.split('\n');
              const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
              );
              if (split[1] != '公開日') {
                await RPA.Logger.info(split[1]);
                await RPA.Logger.info(
                  '"公開日"ではないため取得をスキップします'
                );
              } else {
                await RPA.Logger.info(split[1], 'です');
                await YesterdayUpdateTitleList.push(TitleText);
                await RPA.Logger.info('視聴回数　　        →', NumberofViews);
                await YesterdayNumberofViewsList.push(NumberofViews);
              }
            }
            if (split[0] == NewToday && MonthFlag == '今月') {
              await RPA.Logger.info(
                '本日の公開日が一致したタイトルを取得します'
              );
              const Title2: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-list-cell-video video-title-wrapper')[${i}].children[0]`
              );
              const TitleText = await Title2.getText();
              // 5/25 日付の下の文字を取得
              const date: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i}]`
              );
              const JudgeText = await date.getText();
              var split = await JudgeText.split('\n');
              const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
              );
              if (split[1] != '公開日') {
                await RPA.Logger.info(split[1]);
                await RPA.Logger.info(
                  '"公開日"ではないため取得をスキップします'
                );
              } else {
                await RPA.Logger.info(split[1], 'です');
                await TodayUpdateTitleList.push(TitleText);
                await RPA.Logger.info('視聴回数　　        →', NumberofViews);
                await TodayNumberofViewsList.push(NumberofViews);
              }
            }
            if (split[0].includes(TwoMonthsBefore) == true) {
              await RPA.Logger.info('前々月のアップロード日があります');
              LoopFlag = 'false';
              break;
            }
          }
        }
        await RPA.Logger.info('シートのリスト　　　→', TitleList);
        await RPA.Logger.info('ツールのリスト　　　→', YoutubeTitleList);
        await RPA.Logger.info(
          '前日のアップロードのタイトル →',
          YesterdayUpdateTitleList
        );
        await RPA.Logger.info(
          '本日のアップロードのタイトル →',
          TodayUpdateTitleList
        );
        // 5/11 シート内でタイトルが重複している場合、ツール上では最初にヒットしたものを取得するため注意！（ツール上で区別できる要素があればエラー回避可能）
        // 　　  シートにタイトルの記載があっても、ツール上に動画自体がない場合は取得できないため注意！（途中に「ここまで」の文字を入れておけばエラー回避可能）
        for (let i in TitleList) {
          for (let n in YoutubeTitleList) {
            if (YoutubeTitleList[n].indexOf(TitleList[i]) == 0) {
              await RPA.Logger.info('シートのタイトル　　→', TitleList[i]);
              await RPA.Logger.info(
                'ツールのタイトル　　→',
                YoutubeTitleList[n]
              );
              await RPA.Logger.info(`${Number(n) + 1} 番目　一致しました`);
              const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${n}].innerText`
              );
              await RPA.Logger.info('視聴回数　　        →', NumberofViews);
              await NumberofViewsList.push(NumberofViews);
              // ツールと一致したタイトルをリストに再格納
              await MatchTitleList.push(YoutubeTitleList[n]);
              break;
            }
          }
        }
        await RPA.Logger.info('一致したタイトル　　→', MatchTitleList);
        if (NumberofViewsList.length != TitleList.length) {
          if (LoopFlag == 'false') {
            if (TitleList.length != MatchTitleList.length) {
              // 視聴回数を取得できなかったタイトルを元のリストから除外
              for (let i = 0; i <= TitleList.length - 1; i++) {
                if (MatchTitleList.includes(TitleList[i]) == false) {
                  await RPA.Logger.info(
                    `タイトルリストの ${i} 番目のタイトルを削除します`
                  );
                  await TitleList.splice(i, 1);
                }
              }
              await RPA.Logger.info('シートのリスト　　　→', TitleList);
              await RPA.Logger.info('データ取得を終了します');
              await RPA.Logger.info('タイトルリストをクリアします');
              YoutubeTitleList = [];
              break;
            } else {
              await RPA.Logger.info('不一致のタイトルはありませんでした');
              await RPA.Logger.info('データ取得を終了します');
              await RPA.Logger.info('タイトルリストをクリアします');
              YoutubeTitleList = [];
              break;
            }
          } else {
            await RPA.Logger.info(
              '取得が完了していないため次のページに進みます'
            );
            await RPA.Logger.info('タイトルリストをクリアします');
            YoutubeTitleList = [];
            const NextPage = await RPA.WebBrowser.findElementById(
              'navigate-after'
            );
            // 5/13 以下だどなぜかJenkinsでは処理が重複する → ジャバスクのクリックで回避
            // await RPA.WebBrowser.mouseClick(NextPage);
            await NextPage.click();
          }
        } else {
          await RPA.Logger.info('データ取得を終了します');
          await RPA.Logger.info('タイトルリストをクリアします');
          YoutubeTitleList = [];
          await RPA.Logger.info(
            '視聴回数リスト数　　→',
            NumberofViewsList.length
          );
          if (MonthFlag == '今月') {
            await RPA.Logger.info(
              '前日にアップロードされたタイトルの視聴回数の数　→',
              YesterdayNumberofViewsList.length
            );
            await RPA.Logger.info(
              '本日アップロードされたタイトルの視聴回数の数　　→',
              TodayNumberofViewsList.length
            );
          }
          break;
        }
      }
    }
  }
  // 5/11 ページ数を 50 にすると謎のエラーが出る → 30件に固定で回避
  const DeliteIcon = await RPA.WebBrowser.findElementById('delete-icon');
  await RPA.WebBrowser.mouseClick(DeliteIcon);
  await RPA.sleep(1000);
  LoopFlag = 'true';
}

// 5/26 仕様変更
// 一覧が出るまで待機する関数
let LoadingFlag = 'false';
async function Wait() {
  while (LoadingFlag == 'false') {
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
      LoadingFlag = 'true';
      await RPA.Logger.info('＊＊＊動画の一覧が出ました＊＊＊');
    }
  }
  LoadingFlag = 'false';
}

const Column = [];
const ColumnList = [
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
  'CA'
];
// 6/10 追加修正
async function MLeagueSetData() {
  if (YoutubeTitleList.length < 1) {
    await RPA.Logger.info('今月のタイトルが 0 のため記載をスキップします');
  } else {
    await RPA.Logger.info('現在のシートID　　　→', CurrentSSID[0]);
    await RPA.Logger.info(`現在のシート名　　　→`, ProgramName[0]);
    date = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${CurrentSSID[0]}`,
      range: `${ProgramName[0]}!Q3:CA${LastRow2}`
    });
    // 記載する開始列を取得
    for (let i in date[0]) {
      if (Today == date[0][i]) {
        Column[0] = ColumnList[i];
        await RPA.Logger.info(`${Column[0]} 列目に記載を開始します`);
        break;
      }
    }
    await YoutubeTitleList.reverse();
    await UploadDateList.reverse();
    await NumberofViewsList.reverse();
    for (let i = 0; i <= YoutubeTitleList.length - 1; i++) {
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${CurrentSSID[0]}`,
        range: `${ProgramName[0]}!Q${StartRow3 - i}:R${StartRow3 - i}`,
        values: [[YoutubeTitleList[i], UploadDateList[i]]],
        parseValues: true
      });
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${CurrentSSID[0]}`,
        range: `${ProgramName[0]}!${Column[0]}${StartRow3 - i}:${
          Column[0]
        }${StartRow3 - i}`,
        values: [[NumberofViewsList[i]]],
        parseValues: true
      });
      await RPA.Logger.info(`${StartRow3 - i} 行目に記載しました`);
    }
    await RPA.Logger.info(
      `${ProgramName[0]} のタイトルと視聴回数をクリアします`
    );
    YoutubeTitleList = [];
    NumberofViewsList = [];
  }
}

async function SetData() {
  if (
    NumberofViewsList.length < 1 &&
    YesterdayNumberofViewsList.length < 1 &&
    TodayNumberofViewsList.length < 1
  ) {
    await RPA.Logger.info('取得した視聴回数が 0 のため記載をスキップします');
  } else {
    await RPA.Logger.info('現在のシートID　　　→', CurrentSSID[0]);
    await RPA.Logger.info('現在のシート名　　　→', CurrentSSName[0]);
    // 記載する開始列を取得
    for (let i in date[0]) {
      if (Today == date[0][i]) {
        Column[0] = ColumnList[i];
        await RPA.Logger.info(`${Column[0]} 列目に記載を開始します`);
        break;
      }
    }
    // シートにタイトルの記載がない場合は「本数：１」の行から記載
    if (TitleList.length < 1) {
      let Row6 = Row3 - 4;
      await RPA.Logger.info(`${Row6} 行目から記載します`);
      if (MonthFlag == '今月') {
        let Row5;
        // 月初の場合
        if (Today == BeginningOfMonth) {
          await RPA.Logger.info(
            '月初のため、本日アップロードのタイトルの有無のみ確認します'
          );
          if (TodayUpdateTitleList.length < 1) {
            await RPA.Logger.info(
              '本日アップロードされたタイトルが 0 のため記載をスキップします'
            );
          } else {
            await RPA.Logger.info('本日アップロードされたタイトルを記載します');
            await TodayUpdateTitleList.reverse();
            await TodayNumberofViewsList.reverse();
            for (let i = 0; i <= TodayUpdateTitleList.length - 1; i++) {
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${CurrentSSID[0]}`,
                range: `${CurrentSSName[0]}!Q${Row6 - i}:R${Row6 - i}`,
                values: [[TodayUpdateTitleList[i], NewToday]],
                parseValues: true
              });
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${CurrentSSID[0]}`,
                range: `${CurrentSSName[0]}!${Column[0]}${Row6 - i}:${
                  Column[0]
                }${Row6 - i}`,
                values: [[TodayNumberofViewsList[i]]],
                parseValues: true
              });
              await RPA.Logger.info(`${Row6 - i} 行目に記載しました`);
            }
          }
        } else {
          // 両日ともにアップロードされたタイトルがない場合
          if (
            YesterdayUpdateTitleList.length < 1 &&
            TodayUpdateTitleList.length < 1
          ) {
            await RPA.Logger.info(
              'アップロードされたタイトルが 0 のため記載をスキップします'
            );
            // 前日にアップロードされたタイトルがある場合
          } else if (YesterdayUpdateTitleList.length > 0) {
            await RPA.Logger.info(
              '前日にアップロードされたタイトルを記載します'
            );
            // 上に向かって記載するためデータも合わせて逆順にする
            await YesterdayUpdateTitleList.reverse();
            await YesterdayNumberofViewsList.reverse();
            // アップロードされたタイトルとその視聴回数を記載
            for (let i = 0; i <= YesterdayUpdateTitleList.length - 1; i++) {
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${CurrentSSID[0]}`,
                range: `${CurrentSSName[0]}!Q${Row6 - i}:R${Row6 - i}`,
                values: [[YesterdayUpdateTitleList[i], Yesterday]],
                parseValues: true
              });
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${CurrentSSID[0]}`,
                range: `${CurrentSSName[0]}!${Column[0]}${Row6 - i}:${
                  Column[0]
                }${Row6 - i}`,
                values: [[YesterdayNumberofViewsList[i]]],
                parseValues: true
              });
              await RPA.Logger.info(`${Row6 - i} 行目に記載しました`);
            }
            // 前日にアップロードされたタイトルはあるが本日アップロードされたタイトルがない場合
            if (TodayUpdateTitleList.length < 1) {
              await RPA.Logger.info(
                '本日アップロードされたタイトルはありませんでした'
              );
            }
            // 両日ともにアップロードされたタイトルがある場合
            else {
              await RPA.Logger.info(
                '本日アップロードされたタイトルを記載します'
              );
              await TodayUpdateTitleList.reverse();
              await TodayNumberofViewsList.reverse();
              // 前日にアップロードされたタイトルの上に本日アップロードされたタイトルを記載
              for (let i = 0; i <= TodayUpdateTitleList.length - 1; i++) {
                Row5 = YesterdayUpdateTitleList.length;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${CurrentSSID[0]}`,
                  range: `${CurrentSSName[0]}!Q${Row6 - Row5 - i}:R${Row6 -
                    Row5 -
                    i}`,
                  values: [[TodayUpdateTitleList[i], NewToday]],
                  parseValues: true
                });
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${CurrentSSID[0]}`,
                  range: `${CurrentSSName[0]}!${Column[0]}${Row6 - Row5 - i}:${
                    Column[0]
                  }${Row6 - Row5 - i}`,
                  values: [[TodayNumberofViewsList[i]]],
                  parseValues: true
                });
                await RPA.Logger.info(`${Row6 - Row5 - i} 行目に記載しました`);
              }
            }
            // 前日にアップロードされたタイトルはないが本日アップロードされたタイトルがある場合
          } else if (
            YesterdayUpdateTitleList.length < 1 ||
            TodayUpdateTitleList.length > 0
          ) {
            await RPA.Logger.info(
              '本日アップロードされたタイトルのみ記載します'
            );
            await TodayUpdateTitleList.reverse();
            await TodayNumberofViewsList.reverse();
            for (let i = 0; i <= TodayUpdateTitleList.length - 1; i++) {
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${CurrentSSID[0]}`,
                range: `${CurrentSSName[0]}!Q${Row6 - i}:R${Row6 - i}`,
                values: [[TodayUpdateTitleList[i], NewToday]],
                parseValues: true
              });
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${CurrentSSID[0]}`,
                range: `${CurrentSSName[0]}!${Column[0]}${Row6 - i}:${
                  Column[0]
                }${Row6 - i}`,
                values: [[TodayNumberofViewsList[i]]],
                parseValues: true
              });
              await RPA.Logger.info(`${Row6 - i} 行目に記載しました`);
            }
          }
        }
      }
    } else {
      await RPA.Logger.info(`${Row3} 行目から記載します`);
      if (MonthFlag == '今月') {
        let Row5;
        // 5/25 前日・本日ともにシートに記載する仕様に変更
        // 両日ともにアップロードされたタイトルがない場合
        if (
          YesterdayUpdateTitleList.length < 1 &&
          TodayUpdateTitleList.length < 1
        ) {
          await RPA.Logger.info(
            'アップロードされたタイトルが 0 のため記載をスキップします'
          );
          // 前日にアップロードされたタイトルがある場合
        } else if (YesterdayUpdateTitleList.length > 0) {
          await RPA.Logger.info('前日にアップロードされたタイトルを記載します');
          // 上に向かって記載するためデータも合わせて逆順にする
          await YesterdayUpdateTitleList.reverse();
          await YesterdayNumberofViewsList.reverse();
          // 5/12 RPAでは追加で行数を増やすことができないので事前に何行か増やしていただく必要がある
          // アップロードされたタイトルとその視聴回数を記載
          for (let i = 0; i <= YesterdayUpdateTitleList.length - 1; i++) {
            // 5/25 既に同じタイトルが記載されている場合は記載をスキップ
            if (TitleList.includes(YesterdayUpdateTitleList[i]) == true) {
              await RPA.Logger.info(YesterdayUpdateTitleList[i]);
              await RPA.Logger.info(
                ` ↑ 既に同じタイトルが記載されているためスキップします`
              );
            } else {
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${CurrentSSID[0]}`,
                range: `${CurrentSSName[0]}!Q${Row3 - 1 - i}:R${Row3 - 1 - i}`,
                values: [[YesterdayUpdateTitleList[i], Yesterday]],
                parseValues: true
              });
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${CurrentSSID[0]}`,
                range: `${CurrentSSName[0]}!${Column[0]}${Row3 - 1 - i}:${
                  Column[0]
                }${Row3 - 1 - i}`,
                values: [[YesterdayNumberofViewsList[i]]],
                parseValues: true
              });
              await RPA.Logger.info(`${Row3 - 1 - i} 行目に記載しました`);
            }
          }
          // 前日にアップロードされたタイトルはあるが本日アップロードされたタイトルがない場合
          if (TodayUpdateTitleList.length < 1) {
            await RPA.Logger.info(
              '本日アップロードされたタイトルはありませんでした'
            );
          }
          // 両日ともにアップロードされたタイトルがある場合
          else {
            await RPA.Logger.info('本日アップロードされたタイトルを記載します');
            await TodayUpdateTitleList.reverse();
            await TodayNumberofViewsList.reverse();
            // 前日にアップロードされたタイトルの上に本日アップロードされたタイトルを記載
            for (let i = 0; i <= TodayUpdateTitleList.length - 1; i++) {
              if (TitleList.includes(TodayUpdateTitleList[i]) == true) {
                await RPA.Logger.info(TodayUpdateTitleList[i]);
                await RPA.Logger.info(
                  ` ↑ 既に同じタイトルが記載されているためスキップします`
                );
              } else {
                Row5 = YesterdayUpdateTitleList.length;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${CurrentSSID[0]}`,
                  range: `${CurrentSSName[0]}!Q${Row3 - 1 - Row5 - i}:R${Row3 -
                    1 -
                    Row5 -
                    i}`,
                  values: [[TodayUpdateTitleList[i], NewToday]],
                  parseValues: true
                });
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${CurrentSSID[0]}`,
                  range: `${CurrentSSName[0]}!${Column[0]}${Row3 -
                    1 -
                    Row5 -
                    i}:${Column[0]}${Row3 - 1 - Row5 - i}`,
                  values: [[TodayNumberofViewsList[i]]],
                  parseValues: true
                });
                await RPA.Logger.info(
                  `${Row3 - 1 - Row5 - i} 行目に記載しました`
                );
              }
            }
          }
          // 前日にアップロードされたタイトルはないが本日アップロードされたタイトルがある場合
        } else if (
          YesterdayUpdateTitleList.length < 1 ||
          TodayUpdateTitleList.length > 0
        ) {
          await RPA.Logger.info('本日アップロードされたタイトルのみ記載します');
          await TodayUpdateTitleList.reverse();
          await TodayNumberofViewsList.reverse();
          for (let i = 0; i <= TodayUpdateTitleList.length - 1; i++) {
            if (TitleList.includes(TodayUpdateTitleList[i]) == true) {
              await RPA.Logger.info(TodayUpdateTitleList[i]);
              await RPA.Logger.info(
                ` ↑ 既に同じタイトルが記載されているためスキップします`
              );
            } else {
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${CurrentSSID[0]}`,
                range: `${CurrentSSName[0]}!Q${Row3 - 1 - i}:R${Row3 - 1 - i}`,
                values: [[TodayUpdateTitleList[i], NewToday]],
                parseValues: true
              });
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${CurrentSSID[0]}`,
                range: `${CurrentSSName[0]}!${Column[0]}${Row3 - 1 - i}:${
                  Column[0]
                }${Row3 - 1 - i}`,
                values: [[TodayNumberofViewsList[i]]],
                parseValues: true
              });
              await RPA.Logger.info(`${Row3 - 1 - i} 行目に記載しました`);
            }
          }
        }
      }
    }
    // 再度シートのタイトルを取得
    const SheetTitle2 = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${CurrentSSID[0]}`,
      range: `${CurrentSSName[0]}!Q${Row3}:Q${Row4}`
    });
    // 一致したタイトルのみ視聴回数を記載
    for (let i = 0; i <= TitleList.length - 1; i++) {
      for (let n = 0; n <= SheetTitle2.length - 1; n++) {
        if (TitleList[i] == SheetTitle2[n][0]) {
          await RPA.Logger.info('リストのタイトル　　→', TitleList[i]);
          await RPA.Logger.info('シートのタイトル　　→', SheetTitle2[n][0]);
          await RPA.Logger.info('視聴回数　　　　　　→', NumberofViewsList[i]);
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${CurrentSSID[0]}`,
            range: `${CurrentSSName[0]}!${Column[0]}${Row3 + n}:${
              Column[0]
            }${Row3 + n}`,
            values: [[NumberofViewsList[i]]],
            // 数字の前の ' が入力されなくなる
            parseValues: true
          });
        }
      }
    }
    if (
      NumberofViewsList.length < 1 &&
      YesterdayNumberofViewsList.length > 0 &&
      TodayNumberofViewsList.length < 1
    ) {
      await RPA.Logger.info(`【${ProgramName[0]}】の記載をスキップしました`);
    } else {
      await RPA.Logger.info(`【${ProgramName[0]}】の記載が完了しました`);
      await RPA.Logger.info('視聴回数リストをクリアします');
      NumberofViewsList = [];
    }
  }
  await RPA.Logger.info('一致したタイトルをクリアします');
  MatchTitleList = [];
  if (MonthFlag == '今月') {
    await RPA.Logger.info('アップロードされたタイトルと視聴回数をクリアします');
    YesterdayUpdateTitleList = [];
    YesterdayNumberofViewsList = [];
    TodayUpdateTitleList = [];
    TodayNumberofViewsList = [];
  }
}
