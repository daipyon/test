import RPA from 'ts-rpa';
import { WebElement } from 'selenium-webdriver';
const moment = require('moment');

const Today = moment().format('M/D');
RPA.Logger.info(Today);
// 今月をフォーマット変更して取得
const ThisMonth = moment().format('YYYY/MM');
RPA.Logger.info(`今月　　　 → ${ThisMonth}`);
// 前月をフォーマット変更して取得
const LastMonth = moment()
  .subtract(1, 'months')
  .format('YYYY/MM');
RPA.Logger.info(`前月　　　 → ${LastMonth}`);

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [
  `【Youtube 集計】今月シート（Mリーグ） の集計が完了しました`
];

// スプレッドシートID
// 本番用
// const SSID = process.env.Senden_Youtube_SheetID3;

// テスト用
const mySSID = process.env.My_SheetID2;

// 今月のシートIDを取得
let ThisMonthSheetID;
// シート名
const SSName = [`Mリーグ`];

// 作業するスプレッドシートから読み込む行数を記載
let SheetData;
const StartRow = 1;
const LastRow = 200;

// 各番組名の開始行
let ProgramStartRow;
// タイトルの開始行
let TitleStartRow;
// 視聴回数を記載する開始列
let SetShityoukaisuuStartColumn;

// 番組名のリスト
let ProgramList = [];
// 取得したタイトル・アップロード日・視聴回数を格納する
let MLeagueTitleList = [];
let MLeagueShityoukaisuuList = [];

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
    SheetData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: ThisMonthSheetID,
      range: `${SSName}!A${StartRow}:CB${LastRow}`
    });
    // 列（アルファベット）を取得
    await GetColumn();
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
          // N列の4行目を基準にQ列のタイトルがあるかを確認
          ProgramStartRow = Number(d) + 3;
          const Title = await RPA.Google.Spreadsheet.getValues({
            spreadsheetId: ThisMonthSheetID,
            range: `${SSName}!N${ProgramStartRow}:CB${LastRow}`
          });
          // Q列のタイトル数の分だけ回す
          for (let e in Title) {
            TitleStartRow = ProgramStartRow + Number(e);
            if (Title[e][0] == `1`) {
              await RPA.Logger.info(
                `${TitleStartRow} 行目　本数:${Title[e][0]} です.ストップします`
              );
              await Work();
              break;
            }
          }
        }
      }
    }
    // }
    await RPA.Logger.info(Slack_Text[0]);
    // エラー発生時の処理
    if (ErrorText.length >= 1) {
      // const DOM = await RPA.WebBrowser.driver.getPageSource();
      // await RPA.Logger.info(DOM);
      await RPA.SystemLogger.error(ErrorText);
      Slack_Text[0] = `【Youtube 集計】今月シート（Mリーグ） でエラーが発生しました！\n${ErrorText}`;
      await RPA.WebBrowser.takeScreenshot();
    }
    await SlackPost(Slack_Text[0]);
    await RPA.WebBrowser.quit();
    await RPA.sleep(1000);
    await process.exit();
  }
}

Start();

async function Work() {
  try {
    // Youtube Studioにログイン
    await YoutubeLogin();
    // Mリーグのデータを取得
    await GetMLeagueData();
    // 取得したタイトル・視聴回数をシートに記載する
    await SetMLeagueTitle(MLeagueTitleList);
    await SetMLeagueShityoukaisuu(MLeagueShityoukaisuuList);
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
  await RPA.WebBrowser.get(process.env.Youtube_MLeague_Url);
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
      'style-scope ytcp-table-footer'
    );
    if (Filter.length >= 1) {
      await RPA.Logger.info('＊＊＊ログイン成功しました＊＊＊');
      break;
    }
  }
}

// ループを抜けるためのフラグ
let LoopFlag = 'true';
async function GetMLeagueData() {
  while (LoopFlag == 'true') {
    await Wait();
    // ページ遷移できているか確認のため、現在のページ範囲を取得
    const Range = await RPA.WebBrowser.findElementByClassName(
      'page-description style-scope ytcp-table-footer'
    );
    const RangeText = await Range.getText();
    await RPA.Logger.info(`現在のページ範囲　  → ${RangeText}`);
    // 検索にヒットした件数分だけ回す
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
        await RPA.Logger.info(`公開日　：【${split[0]}】`);
        // 視聴回数を取得
        const Shityoukaisu: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('style-scope ytcp-video-row cell-body tablecell-views sortable right-align')[${i}].innerText`
        );
        await RPA.Logger.info(`視聴回数：【${Shityoukaisu}】`);
        // 一度リストに格納
        await MLeagueTitleList.push([Title, split[0]]);
        await MLeagueShityoukaisuuList.push(Shityoukaisu);
      }
      // 前月の日付を取得した時点でループを抜ける
      if (split[0].includes(LastMonth) == true) {
        LoopFlag = 'false';
        await RPA.Logger.info(`前月アップロードのタイトルです.ストップします`);
        await RPA.Logger.info(`取得した合計数 　 → ${MLeagueTitleList.length}`);
        break;
      }
      // 1ページ目で取得が完了しなければ次のページに進む
      if (i == 29) {
        await RPA.Logger.info('取得が完了していないため、次のページに進みます');
        const NextPage = await RPA.WebBrowser.findElementById('navigate-after');
        await NextPage.click();
        await RPA.sleep(3000);
        break;
      }
    }
  }
  // 逆順に記載するためリストを反転させる
  await MLeagueTitleList.reverse();
  await MLeagueShityoukaisuuList.reverse();
}

async function SetMLeagueTitle(MLeagueTitleList) {
  // N列の本数:1 の行から逆順に記載する
  for (let i = 0; i <= MLeagueTitleList.length - 1; i++) {
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: ThisMonthSheetID,
      range: `${SSName}!Q${TitleStartRow - i}:R${TitleStartRow - i}`,
      values: [[MLeagueTitleList[i][0], MLeagueTitleList[i][1]]],
      // 数字の前の ' が入力されなくなる
      parseValues: true
    });
  }
}

async function SetMLeagueShityoukaisuu(MLeagueShityoukaisuuList) {
  // N列の本数:1 の行から逆順に記載する
  for (let i = 0; i <= MLeagueShityoukaisuuList.length - 1; i++) {
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: ThisMonthSheetID,
      range: `${SSName}!${SetShityoukaisuuStartColumn}${TitleStartRow -
        i}:${SetShityoukaisuuStartColumn}${TitleStartRow - i}`,
      values: [[MLeagueShityoukaisuuList[i]]],
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
    if (Today == SheetData[2][i]) {
      SetShityoukaisuuStartColumn = ColumnList[i];
      break;
    }
  }
}
