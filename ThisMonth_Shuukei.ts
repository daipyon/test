import RPA from 'ts-rpa';
import { WebElement } from 'selenium-webdriver';
const moment = require('moment');

const Kyou = moment()
  .subtract(20, 'days')
  .format('M/D');
RPA.Logger.info(Kyou);
const Today = moment()
  .subtract(2, 'days')
  .format('YYYY/MM/DD');
RPA.Logger.info(`今日　　　 → ${Today}`);
const Yesterday = moment()
  .subtract(3, 'days')
  .format('YYYY/MM/DD');
RPA.Logger.info(`前日　　　 → ${Yesterday}`);
// 今月をフォーマット変更して取得
const ThisMonth = moment().format('YYYY/MM');
RPA.Logger.info(`今月　　　 → ${ThisMonth}`);

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Youtube 集計】今月シート の集計が完了しました`];

// スプレッドシートID
// 本番用
// const SSID = process.env.Senden_Youtube_SheetID3;
// テスト
const mySSID = process.env.My_SheetID2;

// 今月のシートIDを取得
let ThisMonthSheetID;
// シート名
const SSName = [`News`, `公式`, `バラエティ`, `恋リア`, `Mリーグ`];
// 現在のシート名
let CurrentSSName;

// 作業するスプレッドシートから読み込む行数を記載
let SheetData;
const StartRow = 1;
const LastRow = 200;

// 各番組名の開始行
let ProgramStartRow;
// タイトルの開始行
let TitleStartRow = 0;
// undefined を格納する行
let UndefinedRow = [];
// 視聴回数を記載する開始行
let SetShityoukaisuuStartRow = [];
// 　↓　シートとツールでタイトルが不一致だと記載がずれるため、最終的にはこの変数で調整している
let NewSetShityoukaisuuStartRow;
// 前日アップロード分の記載開始行
let SetYesterdayUploadDataStartRow;
// 本日アップロード分の記載開始行
let SetTodayUploadDataStartRow;
// Mリーグ分の記載開始行
let SetMLeagueDataStartRow;
// 視聴回数を記載する開始列
let SetShityoukaisuuStartColumn;
let SkipRow = 0;

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
      // 列（アルファベット）を取得
      await GetColumn();
      // B列の番組名をリストに格納
      for (let b in SheetData) {
        await ProgramList.push(SheetData[b][1]);
        // if (SheetData[b][1] == `` || SheetData[b][1] == undefined) {
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
            // N列の4行目を基準にQ列のタイトルがあるかを確認
            ProgramStartRow = Number(d) + 3;
            await RPA.Logger.info(
              `空白の行 → ${SetShityoukaisuuStartColumn} 列: ${ProgramStartRow} 行目`
            );
            const Title = await RPA.Google.Spreadsheet.getValues({
              spreadsheetId: ThisMonthSheetID,
              range: `${SSName[a]}!N${ProgramStartRow}:CB${LastRow}`
            });
            // Q列のタイトル数の分だけ回す
            for (let e in Title) {
              // アップロードされたタイトルがある場合は、元々タイトルが記載されていた行の一つ上から記載する
              NewSetShityoukaisuuStartRow = TitleStartRow + 1;
              await RPA.Logger.info(Title[e][3]);
              if (Title[e][3] != undefined) {
                // Q列のタイトルが一致した行からタイトルの取得をスタート
                TitleStartRow = ProgramStartRow + Number(e);
                await SetShityoukaisuuStartRow.push(TitleStartRow);
                await RPA.Logger.info(
                  `${SetShityoukaisuuStartColumn} 列: ${TitleStartRow} 行目`
                );
                await RPA.Logger.info(Title[e][3]);
                // Q列のタイトルをリストに格納
                await TitleList.push(Title[e]);
              }
              if (Title[e][3] == undefined) {
                TitleStartRow = ProgramStartRow + Number(e);
                // 空白の行をリストに格納
                await UndefinedRow.push(TitleStartRow);
              }
              // N列の本数が 1 の場合はストップ
              if (Title[e][0] == `1`) {
                await RPA.Logger.info(
                  `↑ 本数: ${Title[e][0]} です.ストップします`
                );
                await RPA.Logger.info(
                  `${SetShityoukaisuuStartColumn} 列: ${NewSetShityoukaisuuStartRow} 行目`
                );
                break;
              }
            }
          }
        }
        // 番組名が一致しない場合はスキップ
        if (CurrentProgramName != undefined) {
          if (TitleList.length > 0) {
            NewSetShityoukaisuuStartRow =
              NewSetShityoukaisuuStartRow - TitleList.length;
            await RPA.Logger.info(
              `${SetShityoukaisuuStartColumn} 列: ${NewSetShityoukaisuuStartRow +
                1} 行目から開始します`
            );
          }
          if (TitleList.length == 0) {
            await RPA.Logger.info(
              `${SetShityoukaisuuStartColumn} 列: ${NewSetShityoukaisuuStartRow} 行目から開始します`
            );
          }
          await RPA.Logger.info(
            `番組名：【${CurrentProgramName}】タイトル数 → ${TitleList.length}`
          );
          await Work();
          await RPA.Logger.info(`タイトル・視聴回数リストをクリアしました`);
          // 格納したタイトルリストを空にする
          TitleList = [];
          SetShityoukaisuuStartRow = [];
          Count = 0;
          SkipRow = 0;
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
let NotitleFlag = `false`;
async function Work() {
  try {
    if (CurrentSSName == `Mリーグ`) {
      await RPA.Logger.info(`【${CurrentSSName}】アカウントに切り替えます`);
      await YoutubeLogout();
      await MLeagueStart();
      await MLeagueGetData();
    } else {
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
        CurrentProgramName == `スポット` ||
        CurrentProgramName == `今日好き` ||
        CurrentProgramName == `恋リアスポット`
      ) {
        await Spot();
      } else {
        // 番組名を入力
        await SetProgram();
        // 一致する動画がない場合はスキップ
        await NoTitle();
        if (NotitleFlag == `false`) {
          // 前日にアップロードされたデータを番組名から取得
          await GetYesterdayData();
        }
        // フラグを元に戻す
        NotitleFlag = `false`;
      }
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

let Icon;
let UploadVideo;
async function SetProgram() {
  await RPA.Logger.info(`現在の番組名：【${CurrentProgramName}】`);
  await Wait();
  // アイコンをクリック
  Icon = await RPA.WebBrowser.findElementById('text-input');
  await RPA.WebBrowser.mouseClick(Icon);
  await RPA.sleep(2000);
  await RPA.WebBrowser.takeScreenshot();
  // 「タイトル」をクリック
  const Title = await RPA.WebBrowser.findElementById('text-item-0');
  await RPA.WebBrowser.mouseClick(Title);
  await RPA.sleep(1000);
  await RPA.WebBrowser.takeScreenshot();
  // 格納した番組名を入力
  const InputTitle: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByTagName('input')[2]`
  );
  // 格納した番組名を入力
  await RPA.WebBrowser.sendKeys(InputTitle, [CurrentProgramName]);
  await RPA.sleep(1000);
  await RPA.WebBrowser.takeScreenshot();
  // 「適用」をクリック
  const Application = await RPA.WebBrowser.findElementById('apply-button');
  await RPA.WebBrowser.mouseClick(Application);
  await RPA.sleep(3000);
  await RPA.WebBrowser.takeScreenshot();
  // カーソルが被って視聴回数が取得できないため、画面上部にカーソルを動かす
  UploadVideo = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      id: 'video-list-uploads-tab'
    }),
    15000
  );
  await RPA.WebBrowser.mouseMove(UploadVideo);
  await RPA.sleep(1000);
  await RPA.WebBrowser.takeScreenshot();
}

let ShityouKaisuu;
let Total;
let YesterdayTitleList = [];
let YesterdayShityoukaisuuList = [];
async function GetYesterdayData() {
  await Wait();
  await RPA.Logger.info(`前日にアップロードされたタイトルがあるか確認します`);
  // ページ遷移できているか確認のため、現在のページ範囲を取得
  const Range = await RPA.WebBrowser.findElementByClassName(
    'page-description style-scope ytcp-table-footer'
  );
  const RangeText = await Range.getText();
  var text = await RangeText.split(/[\～\/\s]+/);
  Total = Number(text[1]);
  await RPA.Logger.info(`現在のページ範囲　  → ${RangeText}`);
  // １ページ分だけ回す
  for (let i = 1; i <= Total; i++) {
    // 日付を取得
    const UpdateDate: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i -
        1}]`
    );
    const UpdateDateText = await UpdateDate.getText();
    var split = await UpdateDateText.split('\n');
    if (split[0] == Today) {
      await RPA.Logger.info(
        `日付：【${split[0]}】 本日アップロードされたタイトルのためスキップします`
      );
    }
    if (split[0] == Yesterday && split[1] == '公開日') {
      await RPA.Logger.info(`${split[1]} です.取得します`);
      // ツールのタイトルを取得
      const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${i -
          1}].children[0].children[0].children[0].getAttribute('aria-label')`
      );
      await RPA.Logger.info(`タイトル：【${Title}】`);
      // 視聴回数を取得
      const Shityoukaisu: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i -
          1}].innerText`
      );
      await RPA.Logger.info(`視聴回数：【${Shityoukaisu}】`);
      // 一度、前日分リストに格納
      await YesterdayTitleList.push([Title, Yesterday]);
      await YesterdayShityoukaisuuList.push(Shityoukaisu);
    }
  }
  // await YesterdayTitleList.reverse();
  if (TitleList.length == 0) {
    await YesterdayTitleList.reverse();
    // 取得したタイトル・視聴回数をシートに記載する（前日分）
    await SetYesterdayTitleList(YesterdayTitleList);
    await SetYesterdayShityoukaisuuList(YesterdayShityoukaisuuList);
  } else {
    // 既に同じタイトルが記載されていないか確認する
    for (let i in YesterdayTitleList) {
      // for (let n in TitleList) {
      // 記載されている場合はスキップ
      if (TitleList[i][3] == YesterdayTitleList[i][0]) {
        await RPA.Logger.info(YesterdayTitleList[i]);
        await RPA.Logger.info(
          ` ↑ 既に同じタイトルが記載されているためスキップします`
        );
        // 記載されている場合は記載する行を詰める
        SkipRow = Number(i) + 1;
        // break;
      } else {
        await YesterdayTitleList.reverse();
        // 取得したタイトル・視聴回数をシートに記載する（前日分）
        await SetYesterdayTitleList(YesterdayTitleList);
        await SetYesterdayShityoukaisuuList(YesterdayShityoukaisuuList);
        // break;
        // 判定のために元に戻す
        await YesterdayTitleList.reverse();
      }
      // }
    }
  }
  // 本日アップロードされたデータを番組名から取得
  await GetTodayData();
}

let TodayTitleList = [];
let TodayShityoukaisuuList = [];
async function GetTodayData() {
  await Wait();
  await RPA.Logger.info(`本日アップロードされたタイトルがあるか確認します`);
  // ページ遷移できているか確認のため、現在のページ範囲を取得
  const Range = await RPA.WebBrowser.findElementByClassName(
    'page-description style-scope ytcp-table-footer'
  );
  const RangeText = await Range.getText();
  var text = await RangeText.split(/[\～\/\s]+/);
  Total = Number(text[1]);
  await RPA.Logger.info(`現在のページ範囲　  → ${RangeText}`);
  // １ページ分だけ回す
  for (let i = 1; i <= Total; i++) {
    // 日付を取得
    const UpdateDate: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-date sortable column-sorted')[${i -
        1}]`
    );
    const UpdateDateText = await UpdateDate.getText();
    var split = await UpdateDateText.split('\n');
    if (split[0] == Yesterday) {
      await RPA.Logger.info(
        `日付：【${split[0]}】 前日アップロードされたタイトルのためスキップします`
      );
    }
    if (split[0] == Today && split[1] == '公開日') {
      // ツールのタイトルを取得
      const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${i -
          1}].children[0].children[0].children[0].getAttribute('aria-label')`
      );
      await RPA.Logger.info(`タイトル：【${Title}】`);
      // 視聴回数を取得
      const Shityoukaisuu: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i -
          1}].innerText`
      );
      await RPA.Logger.info(`視聴回数：【${Shityoukaisuu}】`);
      // 一度リストに格納
      await TodayTitleList.push([Title, Today]);
      await TodayShityoukaisuuList.push(Shityoukaisuu);
    }
  }
  // await TodayTitleList.reverse();
  if (TitleList.length == 0) {
    await TodayTitleList.reverse();
    // 取得したタイトル・視聴回数をシートに記載する（前日分）
    await SetTodayTitleList(TodayTitleList);
    await SetTodayShityoukaisuuList(TodayShityoukaisuuList);
  } else {
    // 既に同じタイトルが記載されていないか確認する
    for (let i in TodayTitleList) {
      // for (let n in TitleList) {
      // 記載されている場合はスキップ
      if (TitleList[i][3] == TodayTitleList[i][0]) {
        await RPA.Logger.info(TodayTitleList[i]);
        await RPA.Logger.info(
          ` ↑ 既に同じタイトルが記載されているためスキップします`
        );
        // 記載されている場合は記載する行を詰める
        SkipRow = Number(i) + 1;
        // break;
      } else {
        await TodayTitleList.reverse();
        // 取得したタイトル・視聴回数をシートに記載する（本日分）
        await SetTodayTitleList(TodayTitleList);
        await SetTodayShityoukaisuuList(TodayShityoukaisuuList);
        // break;
        // 判定のために元に戻す
        await TodayTitleList.reverse();
      }
      // }
    }
  }
  // 記載したら前日分・本日分リストを空にする
  YesterdayTitleList = [];
  YesterdayShityoukaisuuList = [];
  TodayTitleList = [];
  TodayShityoukaisuuList = [];
  // 元々シートに記載されていたタイトルの視聴回数のみを取得
  await GetShityoukaisuu();
}

let LoopCount = 0;
let MatchCount = 0;
let FisrtGetDataFlag = `true`;
async function GetShityoukaisuu() {
  // 元々シートに記載されていたタイトルがなければスキップ
  if (TitleList.length == 0) {
    await RPA.Logger.info(
      `番組名：【${CurrentProgramName}】タイトルがないためスキップします`
    );
  } else {
    await RPA.Logger.info(`タイトルリスト　　　→ ${TitleList.length} 個`);
    // ページ遷移できているか確認のため、現在のページ範囲を取得
    const Range = await RPA.WebBrowser.findElementByClassName(
      'page-description style-scope ytcp-table-footer'
    );
    const RangeText = await Range.getText();
    var text = await RangeText.split(/[\～\/\s]+/);
    Total = Number(text[1]);
    await RPA.Logger.info(`現在のページ範囲　  → ${RangeText}`);
    // 検索にヒットしたタイトルが30件未満の場合
    if (Total < 30) {
      await RPA.Logger.info(`ヒットした件数：${Total} 件 → 30件未満です`);
      await GetShityoukaisuu2();
      // 検索にヒットしたタイトルが30件以上の場合
    } else {
      // 3ページ分だけ回す
      while (LoopCount < 2) {
        await Wait();
        // 2ページ目以降は残りの検索件数の結果によって処理が分かれる
        if (FisrtGetDataFlag == `false`) {
          Total = Number(text[3]) - 30;
          await RPA.Logger.info(`残りの検索件数： ${Total} 件`);
        }
        // シートのタイトル数の分だけ回す
        for (let i in TitleList) {
          // １ページ分 = 30件分だけ回す
          for (let n = 0; n <= 29; n++) {
            // 2ページ目で検索結果が30件以下の場合
            if (FisrtGetDataFlag == `false` && Total <= 30) {
              await RPA.Logger.info(
                `ヒットした件数：${Total} 件 → 30件以下です`
              );
              await GetShityoukaisuu2();
              break;
              // 検索にヒットしたタイトルが30件以上の場合
            } else {
              // ツールのタイトルを取得
              const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
                `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${n}].children[0].children[0].children[0].getAttribute('aria-label')`
              );
              if (Title == TitleList[i][3]) {
                NewSetShityoukaisuuStartRow =
                  SetShityoukaisuuStartRow[0] + Count;
                await RPA.Logger.info(`タイトル：【${Title}】 一致しました`);
                // 視聴回数を取得
                const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
                  `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${n}].innerText`
                );
                await RPA.Logger.info(`視聴回数：【${NumberofViews}】`);
                ShityouKaisuu = NumberofViews;
                // 取得した視聴回数をシートに記載する
                await SetShityoukaisuu(ShityouKaisuu);
                MatchCount++;
                break;
              }
            }
          }
          Count++;
        }
        // 取得した視聴回数の数 = タイトルの数が一致の場合はストップ
        if (MatchCount == TitleList.length) {
          break;
        }
        // 取得が完了していない場合は次のページに進む
        if (MatchCount != TitleList.length) {
          FisrtGetDataFlag == 'false';
          await RPA.Logger.info(
            '視聴回数の取得が完了していないため次のページに進みます'
          );
          const NextPage = await RPA.WebBrowser.findElementById(
            'navigate-after'
          );
          await NextPage.click();
          await Wait();
          LoopCount++;
          Count = 0;
        }
      }
    }
  }
  await RPA.Logger.info(`取得した視聴回数の数 → ${MatchCount}`);
  await RPA.Logger.info('視聴回数の取得が完了しました');
  const DeliteIcon = await RPA.WebBrowser.findElementById('delete-icon');
  await RPA.WebBrowser.mouseClick(DeliteIcon);
  await RPA.sleep(1000);
  LoopCount = 0;
  MatchCount = 0;
}

async function GetShityoukaisuu2() {
  // 元々シートに記載されていたタイトルがなければスキップ
  for (let i in TitleList) {
    for (let n = 0; n <= Total - 1; n++) {
      // タイトルを取得
      const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${n}].children[0].children[0].children[0].getAttribute('aria-label')`
      );
      if (Title == TitleList[i][3]) {
        NewSetShityoukaisuuStartRow = SetShityoukaisuuStartRow[0] + Count;
        await RPA.Logger.info(`タイトル：【${Title}】一致しました`);
        // 視聴回数を取得
        const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${n}].innerText`
        );
        await RPA.Logger.info(`視聴回数　　        → ${NumberofViews}`);
        ShityouKaisuu = NumberofViews;
        await SetShityoukaisuu(ShityouKaisuu);
        MatchCount++;
        break;
      }
    }
    Count++;
  }
}

// 直接タイトルを入力して視聴回数を取得する
async function Spot() {
  // 元々シートに記載されていたタイトルがなければスキップ
  if (TitleList.length == 0) {
    await RPA.Logger.info(
      `番組名：【${CurrentProgramName}】タイトルがないためスキップします`
    );
  } else {
    await RPA.Logger.info(`タイトルごとに視聴回数を取得します`);
    await RPA.Logger.info(`タイトルリスト　　　→ ${TitleList.length} 個`);
    for (let i in TitleList) {
      await Wait();
      // アイコンをクリック
      Icon = await RPA.WebBrowser.findElementById('text-input');
      await RPA.WebBrowser.mouseClick(Icon);
      await RPA.sleep(2000);
      // 「タイトル」をクリック
      const Title = await RPA.WebBrowser.findElementById('text-item-0');
      await RPA.WebBrowser.mouseClick(Title);
      await RPA.sleep(1000);
      // 格納したタイトルを入力
      const InputTitle: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByTagName('input')[2]`
      );
      await RPA.Logger.info(`シートのタイトル　 → ${TitleList[i][3]}`);
      await RPA.WebBrowser.sendKeys(InputTitle, [TitleList[i][3]]);
      await RPA.sleep(1000);
      // 「適用」をクリック
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
      await NoTitle();
      if (NotitleFlag == `true`) {
        // 一致しない場合は該当行を空白にする
        NewSetShityoukaisuuStartRow = SetShityoukaisuuStartRow[0] + Count;
      }
      if (NotitleFlag == `false`) {
        // 視聴回数を取得
        const NumberofViews: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[0].innerText`
        );
        await RPA.Logger.info(`視聴回数：【${NumberofViews}】`);
        ShityouKaisuu = NumberofViews;
        NewSetShityoukaisuuStartRow = SetShityoukaisuuStartRow[0] + Count;
        // 取得した視聴回数をシートに記載
        await SetShityoukaisuu(ShityouKaisuu);
        const DeliteIcon = await RPA.WebBrowser.findElementById('delete-icon');
        await RPA.WebBrowser.mouseClick(DeliteIcon);
        await RPA.sleep(1000);
      }
      // 都度フラグを元に戻す
      NotitleFlag = `false`;
      Count++;
    }
  }
  await RPA.Logger.info('視聴回数の取得を終了します');
}

async function YoutubeLogout() {
  const AccountButton = await RPA.WebBrowser.findElementById(`account-button`);
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
}

async function MLeagueStart() {
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
  await RPA.sleep(2000);

  // ヘッドレスモードオフ（テスト）用
  // const LoginID = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({ id: 'identifierId' }),
  //   8000
  // );
  // await RPA.WebBrowser.sendKeys(LoginID, [
  //   `${process.env.Youtube_Login_MLeagueID}`
  // ]);
  // const NextButton1 = await RPA.WebBrowser.findElementById('identifierNext');
  // await RPA.WebBrowser.mouseClick(NextButton1);
  // await RPA.sleep(5000);
  // const LoginPW = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({ name: 'password' }),
  //   8000
  // );
  // await RPA.WebBrowser.sendKeys(LoginPW, [
  //   `${process.env.Youtube_Login_MLeaguePW}`
  // ]);
  // const NextButton2 = await RPA.WebBrowser.findElementById('passwordNext');
  // await RPA.WebBrowser.mouseClick(NextButton2);

  // 本番・ヘッドレスモードオン（テスト）用
  const LoginID = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'Email' }),
    8000
  );
  await RPA.WebBrowser.sendKeys(LoginID, [
    `${process.env.Youtube_Login_MLeagueID}`
  ]);
  const NextButton1 = await RPA.WebBrowser.findElementById('next');
  await RPA.WebBrowser.mouseClick(NextButton1);
  await RPA.sleep(5000);
  const GoogleLoginPW_1 = await RPA.WebBrowser.findElementsById('Passwd');
  const GoogleLoginPW_2 = await RPA.WebBrowser.findElementsById('password');
  await RPA.Logger.info(`GoogleLoginPW_1 ` + GoogleLoginPW_1.length);
  await RPA.Logger.info(`GoogleLoginPW_2 ` + GoogleLoginPW_2.length);
  if (GoogleLoginPW_1.length == 1) {
    await RPA.WebBrowser.sendKeys(GoogleLoginPW_1[0], [
      `${process.env.Youtube_Login_MLeaguePW}`
    ]);
    const NextButton2 = await RPA.WebBrowser.findElementById('signIn');
    await RPA.WebBrowser.mouseClick(NextButton2);
  }
  if (GoogleLoginPW_2.length == 1) {
    const GoogleLoginPW = await RPA.WebBrowser.findElementById('password');
    await RPA.WebBrowser.sendKeys(GoogleLoginPW, [
      `${process.env.Youtube_Login_MLeaguePW}`
    ]);
    const NextButton2 = await RPA.WebBrowser.findElementById('submit');
    await RPA.WebBrowser.mouseClick(NextButton2);
  }
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
  while (0 == 0) {
    await RPA.sleep(5000);
    const Filter = await RPA.WebBrowser.findElementsByClassName(
      'style-scope ytcp-table-footer'
    );
    if (Filter.length >= 1) {
      await RPA.Logger.info(`＊＊＊${CurrentSSName} の管理ページです＊＊＊`);
      break;
    }
  }
  await RPA.WebBrowser.takeScreenshot();
}

let MLeagueTitleList = [];
let MLeagueShityoukaisuuList = [];
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
      await RPA.Logger.info(`${split[1]} です.取得します`);
      // ツールのタイトルを取得
      const Title: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-video floating-column last-floating-column')[${i}].children[0].children[0].children[0].getAttribute('aria-label')`
      );
      await RPA.Logger.info(`タイトル：【${Title}】`);
      // 視聴回数を取得
      const Shityoukaisu: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
      );
      await RPA.Logger.info(`視聴回数：【${Shityoukaisu}】`);
      // 一度、前日分リストに格納
      await MLeagueTitleList.push([Title, Yesterday]);
      await MLeagueShityoukaisuuList.push(Shityoukaisu);
    }
  }
  // 取得したタイトル・視聴回数をシートに記載する（Mリーグ）
  await SetMLeagueTitleList(MLeagueTitleList);
  await SetMLeagueShityoukaisuuList(MLeagueShityoukaisuuList);
}

async function SetYesterdayTitleList(YesterdayTitleList) {
  await RPA.Logger.info(YesterdayTitleList);
  for (let i in YesterdayTitleList) {
    SetYesterdayUploadDataStartRow =
      NewSetShityoukaisuuStartRow - Number(i) + SkipRow;
    await RPA.Logger.info(YesterdayTitleList[i]);
    await RPA.Logger.info(
      `${SetShityoukaisuuStartColumn} 列: ${SetYesterdayUploadDataStartRow} 行目に記載します`
    );
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: ThisMonthSheetID,
      range: `${CurrentSSName}!Q${SetYesterdayUploadDataStartRow}:R${SetYesterdayUploadDataStartRow}`,
      values: [[YesterdayTitleList[i][0], YesterdayTitleList[i][1]]],
      // 数字の前の ' が入力されなくなる
      parseValues: true
    });
  }
}

async function SetYesterdayShityoukaisuuList(YesterdayShityoukaisuuList) {
  await YesterdayShityoukaisuuList.reverse();
  await RPA.Logger.info(YesterdayShityoukaisuuList);
  for (let i in YesterdayShityoukaisuuList) {
    SetYesterdayUploadDataStartRow =
      NewSetShityoukaisuuStartRow - Number(i) + SkipRow;
    await RPA.Logger.info(YesterdayShityoukaisuuList[i]);
    await RPA.Logger.info(
      `${SetShityoukaisuuStartColumn} 列: ${SetYesterdayUploadDataStartRow} 行目に記載します`
    );
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: ThisMonthSheetID,
      range: `${CurrentSSName}!${SetShityoukaisuuStartColumn}${SetYesterdayUploadDataStartRow}:${SetShityoukaisuuStartColumn}${SetYesterdayUploadDataStartRow}`,
      values: [[YesterdayShityoukaisuuList[i]]],
      // 数字の前の ' が入力されなくなる
      parseValues: true
    });
  }
}

async function SetTodayTitleList(TodayTitleList) {
  await RPA.Logger.info(TodayTitleList);
  for (let i in TodayTitleList) {
    SetTodayUploadDataStartRow =
      NewSetShityoukaisuuStartRow -
      Number(i) -
      YesterdayTitleList.length +
      SkipRow;
    await RPA.Logger.info(TodayTitleList[i]);
    await RPA.Logger.info(
      `${SetShityoukaisuuStartColumn} 列: ${SetTodayUploadDataStartRow} 行目に記載します`
    );
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: ThisMonthSheetID,
      range: `${CurrentSSName}!Q${SetTodayUploadDataStartRow}:R${SetTodayUploadDataStartRow}`,
      values: [[TodayTitleList[i][0], TodayTitleList[i][1]]],
      // 数字の前の ' が入力されなくなる
      parseValues: true
    });
  }
}

async function SetTodayShityoukaisuuList(TodayShityoukaisuuList) {
  await TodayShityoukaisuuList.reverse();
  await RPA.Logger.info(TodayShityoukaisuuList);
  for (let i in TodayShityoukaisuuList) {
    SetTodayUploadDataStartRow =
      NewSetShityoukaisuuStartRow -
      Number(i) -
      YesterdayShityoukaisuuList.length +
      SkipRow;
    await RPA.Logger.info(TodayShityoukaisuuList[i]);
    await RPA.Logger.info(
      `${SetShityoukaisuuStartColumn} 列: ${SetTodayUploadDataStartRow} 行目に記載します`
    );
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: ThisMonthSheetID,
      range: `${CurrentSSName}!${SetShityoukaisuuStartColumn}${SetTodayUploadDataStartRow}:${SetShityoukaisuuStartColumn}${SetTodayUploadDataStartRow}`,
      values: [[TodayShityoukaisuuList[i]]],
      // 数字の前の ' が入力されなくなる
      parseValues: true
    });
  }
}

async function SetMLeagueTitleList(MLeagueTitleList) {
  await MLeagueTitleList.reverse();
  await RPA.Logger.info(MLeagueTitleList);
  for (let i in MLeagueTitleList) {
    SetMLeagueDataStartRow = NewSetShityoukaisuuStartRow - Number(i);
    await RPA.Logger.info(MLeagueTitleList[i]);
    await RPA.Logger.info(
      `${SetShityoukaisuuStartColumn} 列: ${SetMLeagueDataStartRow} 行目に記載します`
    );
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: ThisMonthSheetID,
      range: `${CurrentSSName}!Q${SetMLeagueDataStartRow}:R${SetMLeagueDataStartRow}`,
      values: [[MLeagueTitleList[i][0], MLeagueTitleList[i][1]]],
      // 数字の前の ' が入力されなくなる
      parseValues: true
    });
  }
}

async function SetMLeagueShityoukaisuuList(MLeagueShityoukaisuuList) {
  await MLeagueShityoukaisuuList.reverse();
  await RPA.Logger.info(MLeagueShityoukaisuuList);
  for (let i in MLeagueShityoukaisuuList) {
    SetMLeagueDataStartRow = NewSetShityoukaisuuStartRow - Number(i);
    await RPA.Logger.info(MLeagueShityoukaisuuList[i]);
    await RPA.Logger.info(
      `${SetShityoukaisuuStartColumn} 列: ${SetMLeagueDataStartRow} 行目に記載します`
    );
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: ThisMonthSheetID,
      range: `${CurrentSSName}!${SetShityoukaisuuStartColumn}${SetMLeagueDataStartRow}:${SetShityoukaisuuStartColumn}${SetMLeagueDataStartRow}`,
      values: [[MLeagueShityoukaisuuList[i]]],
      // 数字の前の ' が入力されなくなる
      parseValues: true
    });
  }
}

async function SetShityoukaisuu(ShityouKaisuu) {
  await RPA.Logger.info(
    `${SetShityoukaisuuStartColumn} 列: ${NewSetShityoukaisuuStartRow} 行目に記載します`
  );
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: ThisMonthSheetID,
    range: `${CurrentSSName}!${SetShityoukaisuuStartColumn}${NewSetShityoukaisuuStartRow}:${SetShityoukaisuuStartColumn}${NewSetShityoukaisuuStartRow}`,
    values: [[ShityouKaisuu]],
    // 数字の前の ' が入力されなくなる
    parseValues: true
  });
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
      const DeliteIcon = await RPA.WebBrowser.findElementById('delete-icon');
      await RPA.WebBrowser.mouseClick(DeliteIcon);
      await RPA.sleep(1000);
    }
  } catch {}
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
// 記載する開始列（アルファベット）を取得する関数
async function GetColumn() {
  for (let i in ColumnList) {
    if (Kyou == SheetData[2][i]) {
      SetShityoukaisuuStartColumn = ColumnList[i];
      break;
    }
  }
}
