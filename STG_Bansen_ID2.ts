import RPA from "ts-rpa";
import { WebDriver, By, FileDetector, Key } from "selenium-webdriver";
import { rootCertificates } from "tls";
import { worker } from "cluster";
import { cachedDataVersionTag } from "v8";
import { start } from "repl";
import { Command } from "selenium-webdriver/lib/command";
import { drive } from "googleapis/build/src/apis/drive";
// デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
RPA.Logger.level = 'INFO';

// スプレッドシートIDとシート名
const SSID = process.env.Bansen_ID2_SheetID;
const SSName = process.env.Bansen_ID2_SheetName;
const SSName2 = process.env.Bansen_ID2_SheetName2;
// AAAMS(ステージ環境)のログイン用メールアドレス・パスワード
const AAAMSID = process.env.Bansen_ID2_AAAMSID;
const AAAMSPW = process.env.Bansen_ID2_AAAMSPW;

const FirstLoginFlag = ['true'];
async function WorkStart(JudgeCampaignId){
  await RPA.Google.authorize({
    // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: "Bearer",
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });
  
  // キャンペーンIDと広告IDを保持する変数
  const IdList = [['','']];
  // 作業対象行とデータを取得
  const WorkData = [];
  const Row = [];
  await GetDataRow(WorkData, Row);

  // AAAMSにログイン
  if (FirstLoginFlag[0] == 'true') {
    await AAAMSLogin();
  }
  if (FirstLoginFlag[0] == 'false') {
    await AAAMSLogin2();
  }
  // 一度ログインしたら、以降はログインページをスキップ
  FirstLoginFlag[0] = 'false';

  // キャンペーンIDが記載されていなければキャンペーンIDから発番開始
  if (JudgeCampaignId.length < 1) {
    RPA.Logger.info('キャンペーン作成します');
    // キャンペーン作成を押下
    await CampaignStart();

    // 共通のデータを入力
    await CommonData(WorkData);

    // 配信先(F列)、配信方式(G列)に入力されている値から処理を分岐
    const Linear = 'リニア';
    const Video = 'ビデオ,タイムシフト';
    const Ume = '埋め配信';
    const Kpi = 'KPI配信';
    const Shitei = '指定配信';
    // 配信先設定(F列)を判定
    if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Ume) {
      await CampaignDeliver(WorkData);
    } else if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Kpi) {
      await CampaignDeliver(WorkData);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Kpi) {
      await CampaignDeliver2(WorkData);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Shitei) {
      await CampaignDeliver2(WorkData);
      // 以下のパターンの場合は作業をスキップし、スタートに戻る
    } else if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Shitei) {
      RPA.Logger.info(`'配信先が「${Linear}」、配信方式が「${Shitei}」のためスキップします'`);
      const ErrorText = [['エラー', `'配信先が「${Linear}」、配信方式が「${Shitei}」のためスキップしました'`]];
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AH${Row[0]}`,values:ErrorText});
      Start();
    }

    // 共通のエラー判定
//     await CommonJudgeError(Row, WorkData);

    // F列が「リニア」、G列が「埋め配信」の場合
    if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Ume) {
      RPA.Logger.info(`${WorkData[0][0][5]}と${WorkData[0][0][6]}です`);
      RPA.Logger.info('パターン1の処理を開始します');
      // K〜N列、R列を入力
      await Pattern1(Row, WorkData);
    // F列が「リニア」、G列が「KPI配信」の場合
    } else if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Kpi) {
      RPA.Logger.info(`${WorkData[0][0][5]}と${WorkData[0][0][6]}です`);
      RPA.Logger.info('パターン2の処理を開始します');
      // K〜M列、O〜R列を入力
      await Pattern2(Row, WorkData);
    // F列が「ビデオ,タイムシフト」、G列が「KPI配信」の場合
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Kpi) {
      RPA.Logger.info('パターン3の処理を開始します');
      // R列、M列、P列、S列を入力
      await Pattern3(Row, WorkData);
    // F列が「ビデオ,タイムシフト」、G列が「指定配信」の場合
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Shitei) {
      RPA.Logger.info('パターン4の処理を開始します');
      // M列、P列、S〜T列を入力
      await Pattern4(Row, WorkData);
    }

    // クラスタ(M列)を判定
    if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Ume) {
      ;
      // await Cluster(WorkData);
    } else if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Kpi) {
      await Cluster(WorkData);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Kpi) {
      await Cluster2(WorkData);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Shitei) {
      await Cluster2(WorkData);
    }

    // 各パターンのエラー判定
    if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Ume) {
//       await Pattern1JudgeError(WorkData, Row);
      ;
    } else if (WorkData[0][0][5] === Linear && WorkData[0][0][6] === Kpi) {
      await Pattern2JudgeError(WorkData, Row);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Kpi) {
      await Pattern3JudgeError(WorkData, Row);
    } else if (WorkData[0][0][5] === Video && WorkData[0][0][6] === Shitei) {
      await Pattern4JudgeError(WorkData, Row);
    }

    // 発番したキャンペーンIDを取得
    await GetCampaignId(WorkData, IdList, Row);

    // 広告作成を開始
    await AdvertisementStart(WorkData, Row);

    // 広告作成のエラー判定
    // await AdJudgeError(WorkData, Row);

    // 広告作成を発番
    await GetAdvertisementId(IdList, Row);
  }
  // キャンペーンIDが作成されていれば広告作成まで作業をスキップ
  if (JudgeCampaignId.length > 1) {
    RPA.Logger.info('広告作成します');
    await AdvertisementStart2(WorkData, Row);

    // await AdJudgeError2(WorkData, Row);

    await GetAdvertisementId2(Row);
  }
}

async function Start() {
  try {
    await RPA.Google.authorize({
      //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: "Bearer",
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    var count = 0;
    while (0 == 0) {
      // キャンペーン・広告貼り付けシートのA〜BU列を取得
      const JudgeData = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName2}!A2:BU3000`});
      // ID円滑シートのA〜AH列を取得
      const AllData = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName}!A3:AH3000`});
        RPA.Logger.info('キャンペーンIDです');
        RPA.Logger.info(JudgeData[count][65]);
        if (AllData[count][32] == '初期') { 
          const JudgeCampaignId = JudgeData[count][65];
          await WorkStart(JudgeCampaignId);
        }
      count += 1;
      if (AllData[count][32] == '完了')　{
        break;
      }
    }
  } catch (error) {
    RPA.SystemLogger.error(error);
    await RPA.WebBrowser.takeScreenshot();
  }
  await RPA.WebBrowser.quit();
}

Start();

async function AAAMSLogin() {
  await RPA.WebBrowser.get(process.env.AAAMS_Login_URL);
  await RPA.sleep(2000);
  try {
    const AAAMS_loginID_ele = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'//*[@id="auth0-lock-container-1"]/div/div[2]/form/div/div/div[3]/span/div/div/div/div/div/div/div/div/div[4]/div[1]/div/input'}),8000);
    await RPA.WebBrowser.sendKeys(AAAMS_loginID_ele,[AAAMSID]);
    const AAAMS_loginPW_ele = RPA.WebBrowser.findElementByXPath('//*[@id="auth0-lock-container-1"]/div/div[2]/form/div/div/div[3]/span/div/div/div/div/div/div/div/div/div[4]/div[2]/div/div/input');
    await RPA.WebBrowser.sendKeys(AAAMS_loginPW_ele,[AAAMSPW]);
    const AAAMS_LoginNextButton = await RPA.WebBrowser.findElementByXPath('//*[@id="auth0-lock-container-1"]/div/div[2]/form/div/div/button');
    await RPA.WebBrowser.mouseClick(AAAMS_LoginNextButton);
    await RPA.sleep(3000);
  } catch {
    RPA.Logger.info('ログイン画面を飛ばします');
  }
  try {
    const Alartbutton = await RPA.WebBrowser.findElementByXPath('//*[@id="reactroot"]/div/div[5]/div[2]/footer/div[1]');
    await RPA.WebBrowser.mouseClick(Alartbutton);
    await RPA.sleep(2000);
  } catch　{
    RPA.Logger.info('AAAMSアラートが出ませんでしたので次に進みます');
  }
  RPA.Logger.info('番宣アカウントを直接呼び出します');
  await RPA.WebBrowser.get(process.env.AAAMS_Account);
  await RPA.sleep(3000);
  // 更新画面をスルー
  try　{
    const Koushin = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/header/div'}),5000);
    const KoushinText = await Koushin.getText();
    if (KoushinText.length > 1) {
      const NextButton01 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/footer/div[1]');
      await RPA.WebBrowser.mouseClick(NextButton01);
      await RPA.sleep(1000);
    }
  } catch {
    RPA.Logger.info('更新画面は出ませんでした')
  }
}

async function AAAMSLogin2(){
  RPA.Logger.info('番宣アカウントを直接呼び出します')
  await RPA.WebBrowser.get(process.env.AAAMS_Account);
  await RPA.sleep(2300);
}

async function GetDataRow(WorkData, Row) {
  // ID円滑シートの作業用フラグ(AG列)を取得
  const WorkRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG1:AG200`});
  for(let i in WorkRow){
    if (WorkRow[i][0].indexOf('初期') == 0) {
      Row[0] = Number(i) + 1;
      break;
    }
  }
  RPA.Logger.info('この行の作業実行します → ',Row[0]);
  // ID円滑シートから作業対象行のデータを取得
  WorkData[0] = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName}!A${Row[0]}:AD${Row[0]}`});
  RPA.Logger.info(WorkData[0]);
  // AG列に”作業中”と記載
  await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:[['作業中']]});
  RPA.Logger.info(`${Row[0]} 行目のステータスを”作業中”に変更しました`);
}

async function CampaignStart() {
  // 左側のキャンペーンをクリック
  const Campaign = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/a[7]');
  await RPA.WebBrowser.mouseClick(Campaign);
  await RPA.sleep(1500);
  // 右上のキャンペーン作成を押下
  const CreateButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[2]/div[3]/div/header/div/div'}),8000);
  await RPA.WebBrowser.mouseClick(CreateButton);
  await RPA.sleep(3000);
}

// 共通のデータを入力
async function CommonData(WorkData) {
  // キャンペーン名(A列)を入力
  const CampaignName = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[2]/div[2]/div/input');
  await RPA.WebBrowser.sendKeys(CampaignName,[WorkData[0][0][0]]);
  // 有効期間をクリック
  const DateRange = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[3]/div[2]/div/div/div[1]/div');
  await RPA.WebBrowser.mouseClick(DateRange);
  // 有効期間：開始(B列)を入力
  const DateRangeStart = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[3]/div[1]/div[1]/input'}), 5000);
  await DateRangeStart.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(DateRangeStart,[WorkData[0][0][1]]);
  await RPA.sleep(300);
  // 有効期間：終了(C列)を入力
  const DateRangeEnd = await RPA.WebBrowser.findElementByXPath('/html/body/div[3]/div[2]/div[1]/input');
  await DateRangeEnd.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(DateRangeEnd,[WorkData[0][0][2]]);
  await RPA.sleep(300);
  // 有効期間のOKボタンを押下
  const OKButton = await RPA.WebBrowser.findElementByXPath('/html/body/div[3]/div[3]/div/button[1]');
  await RPA.WebBrowser.mouseClick(OKButton);
  await RPA.sleep(700);
  // 予約放送枠ID(D列)を入力
  const CampaignSlotId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[1]/div[2]/div/div[1]/input');
  if (WorkData[0][0][3] == 'なし') {
    ;
  } else {
    await RPA.WebBrowser.sendKeys(CampaignSlotId,[WorkData[0][0][3]]);
  }
  // キャンペーン種別(E列)を入力
  const CampaignPromo = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[2]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignPromo,[WorkData[0][0][4]]);
  await RPA.WebBrowser.sendKeys(CampaignPromo,[RPA.WebBrowser.Key.ENTER]);
  // 配信方式設定(G列)を入力
  const CampaignPlacement = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[6]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignPlacement,[WorkData[0][0][6]]);
  await RPA.WebBrowser.sendKeys(CampaignPlacement,[RPA.WebBrowser.Key.ENTER]);
}

// 配信先設定(F列)の判定
async function CampaignDeliver(WorkData) {
  if (WorkData[0][0][5].length < 1) {
    const Uncheck = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[4]/div[2]/div/div[2]/div[1]');
    await RPA.WebBrowser.mouseClick(Uncheck);
  } else {
    ;
  }
}

async function CampaignDeliver2(WorkData) {
  if (WorkData[0][0][5].length < 1) {
    const Uncheck = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[7]/div[2]/div/div[2]/div[1]');
    await RPA.WebBrowser.mouseClick(Uncheck);
  } else {
    const Uncheck = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[7]/div[2]/div/div[2]/div[1]');
    await RPA.WebBrowser.mouseClick(Uncheck);
    const Check = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[7]/div[2]/div/div[2]/div[2]');
    await RPA.WebBrowser.mouseClick(Check);
    const Check2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[7]/div[2]/div/div[2]/div[3]');
    await RPA.WebBrowser.mouseClick(Check2);
  }
}

// 共通のエラー判定
async function CommonJudgeError(Row, WorkData) {
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    Start();
  }
  const DoubleCampaignName = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[2]/div[1]'}),1000);
  const DoubleCampaignNameText = await DoubleCampaignName.getText();
  if (String(DoubleCampaignNameText) == 'キャンペーン名必須\n同じキャンペーン名が既に存在しています') {
    RPA.Logger.info(`エラー【${DoubleCampaignNameText}】`);
    const ErrorText = [['エラー', '同じキャンペーン名が既に存在しています']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AH${Row[0]}`,values:ErrorText});
  }
  if (String(DoubleCampaignNameText) == 'キャンペーン名必須\n最大100文字まで入力できます') {
    RPA.Logger.info(`エラー【${DoubleCampaignNameText}】`);
    const ErrorText2 = [['エラー', 'キャンペーン名が100文字以上です']];
    // const ErrorText2 = [['エラー', 'キャンペーン名が65文字以上です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AH${Row[0]}`,values:ErrorText2});
  }
  if (String(DoubleCampaignNameText) == 'キャンペーン名必須\n必須入力です') {
    RPA.Logger.info(`エラー【${DoubleCampaignNameText}】`);
    const ErrorText3 = [['エラー', 'キャンペーン名：必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AH${Row[0]}`,values:ErrorText3});
  }
  const error = [['エラー']];
  if (WorkData[0][0][1].length < 1 || WorkData[0][0][2].length < 1 || WorkData[0][0][1].length < 1 && WorkData[0][0][1].length < 1) {
    const ErrorText4 = [['必須項目です']];
    RPA.Logger.info(`エラー【有効期間：${ErrorText4}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AI${Row[0]}:AI${Row[0]}`,values:ErrorText4});
  }
  const NotCampaignSlotId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[2]/div/div[2]'}),1000);
  const NotCampaignSlotIdText = await NotCampaignSlotId.getText();
  if (String(NotCampaignSlotIdText) == '該当する放送枠が存在しません') {
    RPA.Logger.info(`エラー【${NotCampaignSlotIdText}】`);
    const ErrorText5 = [['該当する放送枠が存在しません']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AJ${Row[0]}:AJ${Row[0]}`,values:ErrorText5});
  }
  if (WorkData[0][0][5].length < 1) {
    const ErrorText6 = [['必須入力です']];
    RPA.Logger.info(`エラー【${ErrorText6}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AK${Row[0]}:AK${Row[0]}`,values:ErrorText6});
  }
  if (WorkData[0][0][6].length < 1) {
    const ErrorText7 = [['必須入力です']];
    RPA.Logger.info(`エラー【${ErrorText7}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AL${Row[0]}:AL${Row[0]}`,values:ErrorText7});
  }
}

// H〜L列、N列、R列を入力
async function Pattern1(Row, WorkData) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    Start();
  }
  // 配信先曜日×時間帯テンプレート(H列)を入力
  const CampaignDayHourTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[2]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[WorkData[0][0][7]]);
  await RPA.sleep(1000);
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[RPA.WebBrowser.Key.ENTER]);
  // 配信先フィルタテンプレート(I列)を入力
  const CampaignFilterTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[3]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[WorkData[0][0][8]]);
  await RPA.sleep(1000);
  await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[RPA.WebBrowser.Key.ENTER]);
  // 選択肢が複数出た場合は指定したものを選択
  const FilterTempFlag = [];
  FilterTempFlag[0] = true;
  const FirstFilterTempValue = await RPA.WebBrowser.findElementsByClassName('Select-option');
  const FilterTempValueText = await Promise.all(FirstFilterTempValue.map(async elm => (await elm.getText())));
  RPA.Logger.info('配信先フィルタテンプレート一覧 → '+ FilterTempValueText);
  for (let i in FilterTempValueText) {
    if (WorkData[0][0][8] == FilterTempValueText[i]) {
      RPA.Logger.info("一致しました" + FilterTempValueText[i])
      const FilterTempSelectValue = await RPA.WebBrowser.findElementByXPath(`/html/body/div/div/div[5]/div[2]/div[1]/div/form/div/div[5]/div[2]/div[1]/div/div/div/div[2]/div/div[${Number(i) + Number(1)}]`);
      await RPA.sleep(100);
      await RPA.WebBrowser.mouseClick(FilterTempSelectValue);
      FilterTempFlag[0] = false;
      await RPA.sleep(500);
    }
  }
  // NGシリーズ属性(J列)を入力
  const CampaignAttributeIds = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[4]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[WorkData[0][0][9]]);
  await RPA.sleep(500);
  await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[RPA.WebBrowser.Key.ENTER]);
  // // 隣接許容レベル・広告(K列)を入力
  // if (WorkData[0][0][10].length < 1) {
  //   const CroseButton = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[5]/div[2]/div/div[1]/div[2]/div/div/div/div/div/span[1]/div[1]/span');
  //   await RPA.WebBrowser.mouseClick(CroseButton);
  // } else {
  //   const CampaignAdjacentAcceptableAd = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[1]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
  //   await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableAd,[WorkData[0][0][10]]);
  //   await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableAd,[RPA.WebBrowser.Key.ENTER]);
  // }
  // // 隣接許容レベル・クリエイティブ(L列)を選択
  // if (WorkData[0][0][11].length < 1) {
  //   const CroseButton2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[5]/div[2]/div/div[2]/div[2]/div/div/div/div/div/span[1]/div[1]/span');
  //   await RPA.WebBrowser.mouseClick(CroseButton2);
  // } else {
  //   const CampaignAdjacentAcceptableCreative = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[2]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
  //   await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableCreative,[WorkData[0][0][11]]);
  //   await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableCreative,[RPA.WebBrowser.Key.ENTER]);
  // }
  // 指定シリーズNG(R列)を入力
  const CampaignSeriesNg = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[6]/div[2]/div/div[1]/div[2]/textarea');
  await RPA.WebBrowser.sendKeys(CampaignSeriesNg,[WorkData[0][0][17]]);
  await RPA.WebBrowser.sendKeys(CampaignSeriesNg,[RPA.WebBrowser.Key.ENTER]);
  // 比率(N列)を入力
  const CampaignStockWeight = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[4]/div[7]/div/div[8]/div[2]/div/input');
  await RPA.WebBrowser.sendKeys(CampaignStockWeight,[WorkData[0][0][13]]);
  await RPA.WebBrowser.sendKeys(CampaignStockWeight,[RPA.WebBrowser.Key.ENTER]);
}

// H〜L列、O〜R列を入力
async function Pattern2(Row, WorkData) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    Start();
  }
  // 配信先曜日×時間帯テンプレート(H列)を入力
  const CampaignDayHourTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[2]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[WorkData[0][0][7]]);
  await RPA.sleep(1000);
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[RPA.WebBrowser.Key.ENTER]);
  // 配信先フィルタテンプレート(I列)を入力
  const CampaignFilterTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[WorkData[0][0][8]]);
  await RPA.sleep(1000);
  // 選択肢が複数出た場合は指定したものを選択
  const FilterTempFlag = [];
  FilterTempFlag[0] = true;
  const FirstFilterTempValue = await RPA.WebBrowser.findElementsByClassName('Select-option');
  const FilterTempValueText = await Promise.all(FirstFilterTempValue.map(async elm => (await elm.getText())));
  RPA.Logger.info('配信先フィルタテンプレート一覧 → '+ FilterTempValueText);
  for (let i in FilterTempValueText) {
    if (WorkData[0][0][8] == FilterTempValueText[i]) {
      RPA.Logger.info(`${FilterTempValueText[i]} 一致しました`);
      const FilterTempSelectValue = await RPA.WebBrowser.findElementByXPath(`/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[2]/div[1]/div/div/div/div/div/div[${Number(i) + Number(1)}]`);
      await RPA.sleep(100);
      await RPA.WebBrowser.mouseClick(FilterTempSelectValue);
      FilterTempFlag[0] = false;
      await RPA.sleep(500);
    }
  }
  // NGシリーズ属性(J列)を入力
  const CampaignAttributeIds = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[4]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[WorkData[0][0][9]]);
  await RPA.sleep(500);
  await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[RPA.WebBrowser.Key.ENTER]);
  // 隣接許容レベル・広告(K列)を入力
  if (WorkData[0][0][10].length < 1) {
    const CroseButton = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[1]/div[2]/div/div/div/div/div/span[2]');
    await RPA.WebBrowser.mouseClick(CroseButton);
  } else {
    const CampaignAdjacentAcceptableAd = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[1]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableAd,[WorkData[0][0][10]]);
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableAd,[RPA.WebBrowser.Key.ENTER]);
  }
  // 隣接許容レベル・クリエイティブ(L列)を選択
  if (WorkData[0][0][11].length < 1) {
    const CroseButton2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[2]/div[2]/div/div/div/div/div/span[2]');
    await RPA.WebBrowser.mouseClick(CroseButton2);
  } else {
    const CampaignAdjacentAcceptableCreative = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[2]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableCreative,[WorkData[0][0][11]]);
    await RPA.WebBrowser.sendKeys(CampaignAdjacentAcceptableCreative,[RPA.WebBrowser.Key.ENTER]);
  }
  // 指定シリーズNG(R列)を入力
  const CampaignSeriesNg = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[1]/div[2]/textarea');
  await RPA.WebBrowser.sendKeys(CampaignSeriesNg,[WorkData[0][0][17]]);
  await RPA.WebBrowser.sendKeys(CampaignSeriesNg,[RPA.WebBrowser.Key.ENTER]);
  // 優先度(O列)を入力
  const GoalPriority = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[8]/div[2]/div/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(GoalPriority,[WorkData[0][0][14]]);
  await RPA.WebBrowser.sendKeys(GoalPriority,[RPA.WebBrowser.Key.ENTER]);
  // 総目標imp(P列)を入力
  const GoalGimp = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[9]/div[2]/div/input');
  await RPA.WebBrowser.sendKeys(GoalGimp,[WorkData[0][0][15]]);
  await RPA.WebBrowser.sendKeys(GoalGimp,[RPA.WebBrowser.Key.ENTER]);
  if (WorkData[0][0][16].length < 1) {
    ;
  } else {
    // セルの値を期間目標impと目標期間に分割
    const ValuesQ = WorkData[0][0][16].split('\n');
    // さらに目標期間を開始時間と終了時間に分割
    const ValuesQ2 = ValuesQ[1].split('-');
    // フォーマットを揃えるため、今日の日付から「年」を取得
    var dd = new Date();
    var yyyy = dd.getFullYear();
    // 一番前に「年」を、一番後ろに「秒」をそれぞれ追加
    const start = `${yyyy}/` + `${ValuesQ2[0]}` + ':00';
    const end = `${yyyy}/` + `${ValuesQ2[1]}` + ':00';
    // プラスボタンを押下
    const PlusButton = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[12]/div[3]/div');
    await RPA.WebBrowser.mouseClick(PlusButton);
    await RPA.sleep(500);
    // 目標期間をクリック
    const DateRange = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[12]/div[2]/div/div[1]/div[1]/div[2]/div/div/div[1]/div');
    await RPA.WebBrowser.mouseClick(DateRange);
    // 目標期間：開始を入力
    const DateRangeStart = await RPA.WebBrowser.findElementByXPath('/html/body/div[4]/div[1]/div[1]/input');
    await DateRangeStart.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(DateRangeStart,[start]);
    await RPA.sleep(300);
    // 目標期間：終了を入力
    const DateRangeEnd = await RPA.WebBrowser.findElementByXPath('/html/body/div[4]/div[2]/div[1]/input');
    await DateRangeEnd.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(DateRangeEnd,[end]);
    await RPA.sleep(300);
    // 目標期間のOKボタンを押下
    const OKButton = await RPA.WebBrowser.findElementByXPath('/html/body/div[4]/div[3]/div/button[1]');
    await RPA.WebBrowser.mouseClick(OKButton);
    await RPA.sleep(700);
    // 期間目標impを入力
    const GoalTermsGimp = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[12]/div[2]/div/div[1]/div[2]/div[2]/div/input');
    await RPA.WebBrowser.sendKeys(GoalTermsGimp,[ValuesQ[0]]);
  }
}

// F列、H〜J列、R列、P列、S列を入力
async function Pattern3(Row, WorkData) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    Start();
  }
  // 配信先曜日×時間帯テンプレート(H列)を入力
  const CampaignDayHourTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[2]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[WorkData[0][0][7]]);
  await RPA.sleep(1000);
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[RPA.WebBrowser.Key.ENTER]);
  // 配信先フィルタテンプレート(I列)を入力
  const CampaignFilterTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[WorkData[0][0][8]]);
  await RPA.sleep(1000);
  // 選択肢が複数出た場合は指定したものを選択
  const FilterTempFlag = [];
  FilterTempFlag[0] = true;
  const FirstFilterTempValue = await RPA.WebBrowser.findElementsByClassName('Select-option');
  const FilterTempValueText = await Promise.all(FirstFilterTempValue.map(async elm => (await elm.getText())));
  RPA.Logger.info('配信先フィルタテンプレート一覧 → '+ FilterTempValueText);
  for (let i in FilterTempValueText) {
    if (WorkData[0][0][8] == FilterTempValueText[i]) {
      RPA.Logger.info(`${FilterTempValueText[i]} 一致しました`);
      const FilterTempSelectValue = await RPA.WebBrowser.findElementByXPath(`/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[2]/div[1]/div/div/div/div/div/div[${Number(i) + Number(1)}]`);
      await RPA.sleep(100);
      await RPA.WebBrowser.mouseClick(FilterTempSelectValue);
      FilterTempFlag[0] = false;
      await RPA.sleep(500);
    }
  }
  // NGシリーズ属性(J列)を入力
  const CampaignAttributeIds = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[4]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[WorkData[0][0][9]]);
  await RPA.sleep(500);
  await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[RPA.WebBrowser.Key.ENTER]);
  // 指定シリーズNG(R列)を入力
  const CampaignSeriesNg = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[1]/div[2]/textarea');
  await RPA.WebBrowser.sendKeys(CampaignSeriesNg,[WorkData[0][0][17]]);
  await RPA.WebBrowser.sendKeys(CampaignSeriesNg,[RPA.WebBrowser.Key.ENTER]);
  // 総目標imp(P列)を入力
  const GoalGimp = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[2]/div/input');
  await RPA.WebBrowser.sendKeys(GoalGimp,[WorkData[0][0][15]]);
  await RPA.WebBrowser.sendKeys(GoalGimp,[RPA.WebBrowser.Key.ENTER]);
  // 訴求シリーズID(S列)を入力
  const CampaignPromoSeriesId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[8]/div[2]/div/div[1]/input');
  await RPA.WebBrowser.sendKeys(CampaignPromoSeriesId,[WorkData[0][0][18]]);
}

// F列、H〜J列、P列、S〜T列を入力
async function Pattern4(Row, WorkData) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    Start();
  }
  // 配信先曜日×時間帯テンプレート(H列)を入力
  const CampaignDayHourTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[2]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[WorkData[0][0][7]]);
  await RPA.sleep(1000);
  await RPA.WebBrowser.sendKeys(CampaignDayHourTempId,[RPA.WebBrowser.Key.ENTER]);
  // 配信先フィルタテンプレート(I列)を入力
  const CampaignFilterTempId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignFilterTempId,[WorkData[0][0][8]]);
  await RPA.sleep(1000);
  // 選択肢が複数出た場合は指定したものを選択
  const FilterTempFlag = [];
  FilterTempFlag[0] = true;
  const FirstFilterTempValue = await RPA.WebBrowser.findElementsByClassName('Select-option');
  const FilterTempValueText = await Promise.all(FirstFilterTempValue.map(async elm => (await elm.getText())));
  RPA.Logger.info('配信先フィルタテンプレート一覧 → '+ FilterTempValueText);
  for (let i in FilterTempValueText) {
    if (WorkData[0][0][8] == FilterTempValueText[i]) {
      RPA.Logger.info(`${FilterTempValueText[i]} 一致しました`);
      const FilterTempSelectValue = await RPA.WebBrowser.findElementByXPath(`/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[2]/div[1]/div/div/div/div/div/div[${Number(i) + Number(1)}]`);
      await RPA.sleep(100);
      await RPA.WebBrowser.mouseClick(FilterTempSelectValue);
      FilterTempFlag[0] = false;
      await RPA.sleep(500);
    }
  }
  // NGシリーズ属性(J列)を入力
  const CampaignAttributeIds = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[4]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input');
  await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[WorkData[0][0][9]]);
  await RPA.sleep(500);
  await RPA.WebBrowser.sendKeys(CampaignAttributeIds,[RPA.WebBrowser.Key.ENTER]);
  // 総目標imp(P列)を入力
  const GoalGimp = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[2]/div/input');
  await RPA.WebBrowser.sendKeys(GoalGimp,[WorkData[0][0][15]]);
  await RPA.WebBrowser.sendKeys(GoalGimp,[RPA.WebBrowser.Key.ENTER]);
  // 訴求シリーズID(S列)を入力
  const CampaignPromoSeriesId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[9]/div[2]/div/div[1]/input');
  await RPA.WebBrowser.sendKeys(CampaignPromoSeriesId,[WorkData[0][0][18]]);
  // 指定シリーズID(T列)を入力
  const CampaignSpecifiedSeriesId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[8]/div[2]/div/div[1]/div[2]/textarea');
  await RPA.WebBrowser.sendKeys(CampaignSpecifiedSeriesId,[WorkData[0][0][19]]);
}

// クラスタ(M列)を判定
async function Cluster(WorkData) {
  // セルの値を分割
  const ValuesM = WorkData[0][0][12].split(',');
  // クラスタ・全て選択をクリック
  const All = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[2]/div/div[1]');
  await RPA.WebBrowser.mouseClick(All);
  for (var i = 0; i <= ValuesM.length - 1; i++) {
    // クラスタ・F1を取得
    const F1 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[2]/div/div[2]/div[1]');
    const CheckF1 = await F1.getText();
    if (CheckF1 == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(F1);
    }
    // クラスタ・F2以上を取得
    const F2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[2]/div/div[2]/div[2]');
    const F2Ijyou = await F2.getText();
    // "以上"という文字を削除
    const CheckF2 = F2Ijyou.slice(0, -2);
    if (CheckF2　== ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(F2);
    }
    // クラスタ・M1を取得
    const M1 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[2]/div/div[2]/div[3]');
    const CheckM1 = await M1.getText();
    if (CheckM1 == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(M1);
    }
    // クラスタ・M2以上を取得
    const M2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[2]/div/div[2]/div[4]');
    const M2Ijyou = await M2.getText();
    // "以上"という文字を削除
    const CheckM2 = M2Ijyou.slice(0, -2);
    if (CheckM2 == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(M2);
    }
    // クラスタ・teenを取得
    const Teen = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[2]/div/div[2]/div[5]');
    const CheckTeen = await Teen.getText();
    if (CheckTeen == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(Teen);
    }
    // クラスタ・otherを取得
    const Other = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[2]/div/div[2]/div[6]');
    const CheckOther = await Other.getText();
    if (CheckOther == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(Other);
    }
    // クラスタ・外部メディア（番宣のみ）を取得
    const Media = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[2]/div/div[2]/div[7]');
    const CheckMedia = await Media.getText();
    if (CheckMedia == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(Media);
    }
    const CheckAll = await All.getText();
    if (CheckAll == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(All);
      // クラスタの外部メデイアのみチェックを外す
      const Uncheck = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[2]/div/div[2]/div[7]');
      await RPA.WebBrowser.mouseClick(Uncheck);
    }
  }
}

async function Cluster2(WorkData) {
  // セルの値を分割
  const ValuesM = WorkData[0][0][12].split(',');
  // クラスタ・全て選択をクリック
  const All = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[1]');
  await RPA.WebBrowser.mouseClick(All);
  for (var i = 0; i <= ValuesM.length - 1; i++) {
    // クラスタ・F1を取得
    const F1 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[2]/div[1]');
    const CheckF1 = await F1.getText();
    if (CheckF1 == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(F1);
    }
    // クラスタ・F2以上を取得
    const F2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[2]/div[2]');
    const F2Ijyou = await F2.getText();
    // "以上"という文字を削除
    const CheckF2 = F2Ijyou.slice(0, -2);
    if (CheckF2　== ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(F2);
    }
    // クラスタ・M1を取得
    const M1 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[2]/div[3]');
    const CheckM1 = await M1.getText();
    if (CheckM1 == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(M1);
    }
    // クラスタ・M2以上を取得
    const M2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[2]/div[4]');
    const M2Ijyou = await M2.getText();
    // "以上"という文字を削除
    const CheckM2 = M2Ijyou.slice(0, -2);
    if (CheckM2 == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(M2);
    }
    // クラスタ・teenを取得
    const Teen = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[2]/div[5]');
    const CheckTeen = await Teen.getText();
    if (CheckTeen == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(Teen);
    }
    // クラスタ・otherを取得
    const Other = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[2]/div[6]');
    const CheckOther = await Other.getText();
    if (CheckOther == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(Other);
    }
    // クラスタ・外部メディア（番宣のみ）を取得
    const Media = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[2]/div[7]');
    const CheckMedia = await Media.getText();
    if (CheckMedia == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(Media);
    }
    const CheckAll = await All.getText();
    if (CheckAll == ValuesM[i]) {
      await RPA.WebBrowser.mouseClick(All);
      // クラスタの外部メデイアのみチェックを外す
      const Uncheck = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[2]/div[7]');
      await RPA.WebBrowser.mouseClick(Uncheck);
    }
  }
}

async function Pattern1JudgeError(WorkData, Row) {
  const error = [['エラー']];
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  const FieldError4 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[2]/div[1]'}),1000);
  const FieldErrorText4 = await FieldError4.getText();
  if (String(FieldErrorText4) == '配信先曜日×時間帯テンプレート必須\n必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText4}】`);
    const ErrorText8 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AM${Row[0]}:AM${Row[0]}`,values:ErrorText8});
  }
  const FieldError5 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[1]'}),1000);
  const FieldErrorText5 = await FieldError5.getText();
  if (String(FieldErrorText5) == 'リニア配信先フィルタテンプレート必須\n必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText5}】`);
    const ErrorText9 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AN${Row[0]}:AN${Row[0]}`,values:ErrorText9});
  }
  const FieldError6 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[1]/div[1]'}),1000);
  const FieldErrorText6 = await FieldError6.getText();
  if (String(FieldErrorText6) == '広告必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText6}】`);
    const ErrorText10 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AO${Row[0]}:AO${Row[0]}`,values:ErrorText10});
  }
  const FieldError7 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[2]/div[1]'}),1000);
  const FieldErrorText7 = await FieldError7.getText();
  if (String(FieldErrorText7) == 'クリエイティブ必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText7}】`);
    const ErrorText11 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AP${Row[0]}:AP${Row[0]}`,values:ErrorText11});
  }
  if (WorkData[0][0][17].length < 1) {
    ;
  } else {
    const NotCampaignSeriesNg = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[2]'}),1000);
    const NotCampaignSeriesNgText = await NotCampaignSeriesNg.getText();
    if (String(NotCampaignSeriesNgText) == 'シリーズIDが存在しません') {
      RPA.Logger.info(`エラー【${NotCampaignSeriesNgText}】`);
      const ErrorText12 = [['シリーズIDが存在しません']];
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AQ${Row[0]}:AQ${Row[0]}`,values:ErrorText12});
    }
  }
  if (WorkData[0][0][12].length < 1) {
    const ErrorText13 = [['必須入力です']];
    RPA.Logger.info(`エラー【クラスタ：${ErrorText13}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AR${Row[0]}:AR${Row[0]}`,values:ErrorText13});
  }
  const FieldError9 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[8]/div[1]'}),1000);
  const FieldErrorText9 = await FieldError9.getText();
    if (String(FieldErrorText9) == '比率必須必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText9}】`);
    const ErrorText14 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AS${Row[0]}:AS${Row[0]}`,values:ErrorText14});
  }
}

async function Pattern2JudgeError(WorkData, Row) {
  const error = [['エラー']];
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  const FieldError4 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[2]/div[1]'}),1000);
  const FieldErrorText4 = await FieldError4.getText();
  if (String(FieldErrorText4) == '配信先曜日×時間帯テンプレート必須\n必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText4}】`);
    const ErrorText8 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AM${Row[0]}:AM${Row[0]}`,values:ErrorText8});
  }
  const FieldError5 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[1]'}),1000);
  const FieldErrorText5 = await FieldError5.getText();
  if (String(FieldErrorText5) == 'リニア配信先フィルタテンプレート必須\n必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText5}】`);
    const ErrorText9 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AN${Row[0]}:AN${Row[0]}`,values:ErrorText9});
  }
  const FieldError6 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[1]/div[1]'}),1000);
  const FieldErrorText6 = await FieldError6.getText();
  if (String(FieldErrorText6) == '広告必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText6}】`);
    const ErrorText10 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AO${Row[0]}:AO${Row[0]}`,values:ErrorText10});
  }
  const FieldError7 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[2]/div[1]'}),1000);
  const FieldErrorText7 = await FieldError7.getText();
  if (String(FieldErrorText7) == 'クリエイティブ必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText7}】`);
    const ErrorText11 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AP${Row[0]}:AP${Row[0]}`,values:ErrorText11});
  }
  if (WorkData[0][0][17].length < 1) {
      ;
  } else {
    const NotCampaignSeriesNg = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[6]/div[2]/div/div[2]'}),1000);
    const NotCampaignSeriesNgText = await NotCampaignSeriesNg.getText();
    if (String(NotCampaignSeriesNgText) == 'シリーズIDが存在しません') {
      RPA.Logger.info(`エラー【${NotCampaignSeriesNgText}】`);
      const ErrorText12 = [['シリーズIDが存在しません']];
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AQ${Row[0]}:AQ${Row[0]}`,values:ErrorText12});
    }
  }
  if (WorkData[0][0][12].length < 1) {
    const ErrorText13 = [['必須入力です']];
    RPA.Logger.info(`エラー【クラスタ：${ErrorText13}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AR${Row[0]}:AR${Row[0]}`,values:ErrorText13});
  }
  if (WorkData[0][0][14].length < 1) {
    const ErrorText15 = [['必須入力です']];
    RPA.Logger.info(`エラー【優先度：${ErrorText15}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AT${Row[0]}:AT${Row[0]}`,values:ErrorText15});
  }
  const FieldError11 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[9]/div[1]'}),1000);
  const FieldErrorText11 = await FieldError11.getText();
  if (String(FieldErrorText11) == '総目標imp必須必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText11}】`);
    const ErrorText16 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AU${Row[0]}:AU${Row[0]}`,values:ErrorText16});
  }
}

async function Pattern3JudgeError(WorkData, Row) {
  const error = [['エラー']];
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  const FieldError4 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[2]/div[1]'}),1000);
  const FieldErrorText4 = await FieldError4.getText();
  if (String(FieldErrorText4) == '配信先曜日×時間帯テンプレート必須\n必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText4}】`);
    const ErrorText8 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AM${Row[0]}:AM${Row[0]}`,values:ErrorText8});
  }
  const FieldError5 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[1]'}),1000);
  const FieldErrorText5 = await FieldError5.getText();
  if (String(FieldErrorText5) == 'リニア配信先フィルタテンプレート必須\n必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText5}】`);
    const ErrorText9 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AN${Row[0]}:AN${Row[0]}`,values:ErrorText9});
  }
  if (WorkData[0][0][17].length < 1) {
    ;
  } else {
    const NotCampaignSeriesNg = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[5]/div[2]/div/div[2]'}),1000);
    const NotCampaignSeriesNgText = await NotCampaignSeriesNg.getText();
    if (String(NotCampaignSeriesNgText) == 'シリーズIDが存在しません') {
      RPA.Logger.info(`エラー【${NotCampaignSeriesNgText}】`);
      const ErrorText12 = [['シリーズIDが存在しません']];
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AQ${Row[0]}:AQ${Row[0]}`,values:ErrorText12});
    }
  }
  if (WorkData[0][0][12].length < 1) {
    const ErrorText13 = [['必須入力です']];
    RPA.Logger.info(`エラー【クラスタ：${ErrorText13}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AR${Row[0]}:AR${Row[0]}`,values:ErrorText13});
  }
  const FieldError11 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[1]'}),1000);
  const FieldErrorText11 = await FieldError11.getText();
  RPA.Logger.info(FieldErrorText11);
  if (String(FieldErrorText11) == '総目標imp必須必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText11}】`);
    const ErrorText16 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AU${Row[0]}:AU${Row[0]}`,values:ErrorText16});
  }
  const NotCampaignPromoSeriesId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[8]/div[2]/div/div[2]'}),1000);
  const NotCampaignPromoSeriesIdText = await NotCampaignPromoSeriesId.getText();
  if (String(NotCampaignPromoSeriesIdText) == 'シリーズIDが存在しません') {
    RPA.Logger.info(`エラー【${NotCampaignPromoSeriesIdText}】`);
    const ErrorText18 = [['シリーズIDが存在しません']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AW${Row[0]}:AW${Row[0]}`,values:ErrorText18});
  }
}

async function Pattern4JudgeError(WorkData, Row) {
  const error = [['エラー']];
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  const FieldError4 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[2]/div[1]'}),1000);
  const FieldErrorText4 = await FieldError4.getText();
  if (String(FieldErrorText4) == '配信先曜日×時間帯テンプレート必須\n必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText4}】`);
    const ErrorText8 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AM${Row[0]}:AM${Row[0]}`,values:ErrorText8});
  }
  const FieldError5 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[3]/div[1]'}),1000);
  const FieldErrorText5 = await FieldError5.getText();
  if (String(FieldErrorText5) == 'リニア配信先フィルタテンプレート必須\n必須入力です') {
    RPA.Logger.info(`エラー【${FieldErrorText5}】`);
    const ErrorText9 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AN${Row[0]}:AN${Row[0]}`,values:ErrorText9});
  }
  if (WorkData[0][0][12].length < 1) {
    const ErrorText13 = [['必須入力です']];
    RPA.Logger.info(`エラー【クラスタ：${ErrorText13}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AR${Row[0]}:AR${Row[0]}`,values:ErrorText13});
  }
  const FieldError11 = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[7]/div[1]'}),1000);
  const FieldErrorText11 = await FieldError11.getText();
  RPA.Logger.info(FieldErrorText11);
  if (WorkData[0][0][15].length < 1) {
    RPA.Logger.info(`エラー【${FieldErrorText11}】`);
    const ErrorText16 = [['必須入力です']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AU${Row[0]}:AU${Row[0]}`,values:ErrorText16});
  }
  if (WorkData[0][0][19].length < 1) {
    ;
  } else {
    const NotCampaignSpecifiedSeriesId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[8]/div[2]/div/div[2]'}),1000);
    const NotCampaignSpecifiedSeriesIdText = await NotCampaignSpecifiedSeriesId.getText();
    if (String(NotCampaignSpecifiedSeriesIdText) == 'シリーズIDが存在しません') {
      RPA.Logger.info(`エラー【${NotCampaignSpecifiedSeriesId}】`);
      const ErrorText17 = [['シリーズIDが存在しません']];
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AV${Row[0]}:AV${Row[0]}`,values:ErrorText17});
    }
  }
  const NotCampaignPromoSeriesId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div/div[10]/div/div[9]/div[2]/div/div[2]'}),1000);
  const NotCampaignPromoSeriesIdText = await NotCampaignPromoSeriesId.getText();
  if (String(NotCampaignPromoSeriesIdText) == 'シリーズIDが存在しません') {
    RPA.Logger.info(`エラー【${NotCampaignPromoSeriesIdText}】`);
    const ErrorText18 = [['シリーズIDが存在しません']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AW${Row[0]}:AW${Row[0]}`,values:ErrorText18});
  }
}

// キャンペーンIDを発番
async function GetCampaignId(WorkData, IdList, Row) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    Start();
  }
  // OKボタンをクリック
  const OKButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'//*[@id="reactroot"]/div/div[5]/div[2]/footer/div[2]'}), 5000);
  await RPA.WebBrowser.mouseClick(OKButton);
  await RPA.sleep(5000);
  // キャンペーン名が一致するか判定
  for (var i = 1; i <= 15; i++) {
    const CampaignName = await RPA.WebBrowser.findElementByXPath(`/html/body/div[1]/div/div[2]/div[3]/div/table/tbody/tr[${Number(i)}]/td[3]`);
    const CampaignNameText = await CampaignName.getText();
    RPA.Logger.info('取得したキャンペーン名　　　　　　　　　→　'+CampaignNameText);
    if (CampaignNameText == WorkData[0][0][0]) {
      await RPA.WebBrowser.mouseClick(CampaignName);
      break;
    }
  }
  RPA.Logger.info('現在保持しているデータのキャンペーン名　→　'+WorkData[0][0][0]);
  RPA.Logger.info('キャンペーン名一致、IDを取得します');
  await RPA.sleep(5000);

  // 発番したキャンペーンIDを取得し、キャンペーン・広告貼り付けシートに記載
  const CampaignId = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[2]/div[3]/div/div[1]/div[2]/div[1]/div[2]');
  IdList[0][0] = await CampaignId.getText();
  RPA.Logger.info(IdList);
  // キャンペーン名が同じものがあればその行にもキャンペーンIDを記載
  const JudgeCampaignName = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName2}!AF2:AF3000`});
  for (let i in JudgeCampaignName) {
    if (JudgeCampaignName[i][0] == WorkData[0][0][0]) {
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName2}!BN${Number(i) + 2}:BO${Number(i) + 2}`,values:IdList});
    }
  }
}

async function AdvertisementStart(WorkData, Row) {
  // 右下の広告作成を押下
  const CreateButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div/div/div[2]/div[3]/div/div[2]/div[2]/div'}),8000);
  await RPA.WebBrowser.mouseClick(CreateButton);
  await RPA.sleep(3000);
  // クリエイティブIDを入力
  const CreativeId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[1]/div[2]/div/input');
  await RPA.WebBrowser.sendKeys(CreativeId,[WorkData[0][0][20]]);
  await RPA.sleep(5000);
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  try {
    var ApplyFlag = ['false'];
    await RPA.sleep(300);
    const NotCreativeId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[3]/div[2]/div/div'}),1000);
    const NotCreativeIdText = await NotCreativeId.getText();
    if (String(NotCreativeIdText) == '該当する項目が存在しません') {
      RPA.Logger.info(`エラー【${NotCreativeIdText}】` + ' 作業をスキップします');
      const ErrorText = [['エラー', 'クリエイティブID：該当する項目が存在しません']];
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AH${Row[0]}`,values:ErrorText});
      Start();
    }
  } catch {
    ApplyFlag[0] = 'true';
    RPA.Logger.info('次の処理に進みます');
  }
  // クリエイティブ選択の「選択」をクリック
  const SelectCreative = await RPA.WebBrowser.findElementByXPath('//*[@id="reactroot"]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]/td[1]/div');
  await RPA.WebBrowser.mouseClick(SelectCreative);
  // 配信期間をクリック
  const DateRange = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[2]/div[2]/div/div/div[1]/div');
  await RPA.WebBrowser.mouseClick(DateRange);
  // 配信期間：開始(Z列)を入力
  const DateRangeStart = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[1]/div[1]/input');
  await DateRangeStart.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(DateRangeStart,[WorkData[0][0][25]]);
  await RPA.sleep(300);
  // 配信期間：終了(AA列)を入力
  const DateRangeEnd = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[2]/div[1]/input');
  await DateRangeEnd.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(DateRangeEnd,[WorkData[0][0][26]]);
  await RPA.sleep(300);
  // 配信期間のOKボタンを押下
  const OKButton = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[3]/div/button[1]');
  await RPA.WebBrowser.mouseClick(OKButton);
  // imp比率(AB列)を入力
  if (WorkData[0][0][27].length < 1) {
    const ImpWeight = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[3]/div[2]/div/input');
    await ImpWeight.clear();
  } else {
    const ImpWeight = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[3]/div[2]/div/input');
    await ImpWeight.clear();
    await RPA.WebBrowser.sendKeys(ImpWeight,[WorkData[0][0][27]]);
  }
  // 予約放送枠ID(D列)を入力
  const AdvertisementSlotId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[4]/div[2]/div/div[1]/input');
  await AdvertisementSlotId.clear();
  await RPA.sleep(100);
  if (WorkData[0][0][3] == 'なし') {
    ;
  } else {
    await RPA.WebBrowser.sendKeys(AdvertisementSlotId,[WorkData[0][0][3]]);
  }
  // 視聴開始単価を入力
  const Money = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[7]/div[2]/div/input');
  await RPA.WebBrowser.sendKeys(Money,['100']);
  await RPA.sleep(300);
}

async function AdJudgeError(WorkData, Row) {
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  const error = [['エラー']];
  if (WorkData[0][0][20].length < 1) {
    const ErrorText2 = [['【広告作成】クリエイティブIDが入力されていません']];
    RPA.Logger.info(ErrorText2[0][0]);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AX${Row[0]}:AX${Row[0]}`,values:ErrorText2});
  }
  if (WorkData[0][0][25].length < 1 || WorkData[0][0][26].length < 1 || WorkData[0][0][25].length < 1 && WorkData[0][0][26].length < 1) {
    const ErrorText3 = [['【広告作成】必須項目です']];
    RPA.Logger.info(`エラー【${ErrorText3}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AI${Row[0]}:AI${Row[0]}`,values:ErrorText3});
  }
  if (WorkData[0][0][27].length < 1) {
    const ErrorText4 = [['【広告作成】必須入力です']];
    RPA.Logger.info(`エラー【${ErrorText4}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AS${Row[0]}:AS${Row[0]}`,values:ErrorText4});
  }
  const NotAdvertisementSlotId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[4]/div[2]/div/div[2]'}),1000);
  const NotAdvertisementSlotIdText = await NotAdvertisementSlotId.getText();
  if (String(NotAdvertisementSlotIdText) == '該当する放送枠が存在しません') {
    RPA.Logger.info(`エラー【${NotAdvertisementSlotIdText}】`);
    const ErrorText5 = [['【広告作成】該当する放送枠が存在しません']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AJ${Row[0]}:AJ${Row[0]}`,values:ErrorText5});
  }
}

async function GetAdvertisementId(IdList, Row) {
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    Start();
  }
  // OKボタンをクリック
  const OKButton2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/footer/div[2]');
  await RPA.WebBrowser.mouseClick(OKButton2);
  await RPA.sleep(5000);
  // 発番した広告IDの最右側の「・・・」をマウスオーバー
  const BalloonMenu = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      xpath: '/html/body/div/div/div[2]/div[3]/div/table/tbody/tr/td[9]/div'
    }),
    5000
  );
  await RPA.WebBrowser.mouseMove(BalloonMenu);
  await RPA.sleep(1000);
  // 「配信を変更する」をクリック
  const ChangeDelivery = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[3]/div/div[4]');
  await RPA.WebBrowser.mouseClick(ChangeDelivery);
  await RPA.sleep(1000);
  const OKButton3 = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[5]/div[2]/footer/div[2]');
  await RPA.WebBrowser.mouseClick(OKButton3);
  await RPA.sleep(8000);
  const AdvertisementId = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[2]/div[3]/div/table/tbody/tr/td[1]');
  IdList[0][1] = await AdvertisementId.getText();
  RPA.Logger.info(IdList);
  // キャンペーン・広告貼り付けシートのBN〜BO列に発番した広告IDを記載
  await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName2}!BN${Row[0] - 1}:BO${Row[0] - 1}`,values:IdList});
  // 作業完了を記載
  await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:[['完了']]});
}

async function AdvertisementStart2(WorkData, Row) {
  const LoadingFlag = [];
  LoadingFlag[0] = false;
  while (LoadingFlag[0] == false) {
    const JudgeCampaignId = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName2}!BN${Row[0] - 1}:BN${Row[0] - 1}`});
    // キャンペーンIDで検索
    const SearchByCampaignId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[3]/div/div[1]/form/div/div[5]/div[2]/div/input');
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(SearchByCampaignId, JudgeCampaignId[0]);
    await RPA.sleep(5000);
    try {
      const Loading = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[3]/div/div[2]/div/p[1]');
      const LoadingText = await Loading.getText();
      if (LoadingText == '読み込み中...') {
        LoadingFlag[0] = false;
        RPA.Logger.info('検索結果が出ないため、ブラウザを更新し再検索します')
        await RPA.WebBrowser.refresh();
      }
      RPA.Logger.info('番宣アカウントを直接呼び出します')
      await RPA.WebBrowser.get(process.env.AAAMS_Account);
      await RPA.sleep(3000);
      // 代理店の箇所に「番宣」の文字が入っていない場合は再更新
      const AgencyName = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[3]/div/div[1]/form/div/div[1]/div[2]/div/div[2]/div[2]/div/div/div/div/div/span[1]/div[1]');
      const AgencyNameText = await AgencyName.getText();
      if (AgencyNameText == '代理店名で検索') {
        await RPA.WebBrowser.refresh();
      }
      // 更新画面をスルー
      try {
        const Koushin = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/header/div'}),5000);
        const KoushinText = await Koushin.getText();
        if (KoushinText.length > 1) {
          const NextButton01 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/footer/div[1]');
          await RPA.WebBrowser.mouseClick(NextButton01);
          await RPA.sleep(1000);
        }
      } catch {
        RPA.Logger.info('更新画面は出ませんでした')
      }
      await RPA.sleep(2000);
      const CampaignId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[3]/div/table/tbody/tr[1]/td[3]');
      const CampaignIdText = await CampaignId.getText();
      if (CampaignIdText == JudgeCampaignId[0][0]) {
        LoadingFlag[0] = true;
        RPA.Logger.info('取得したキャンペーンID　　      　 →　'+CampaignIdText);
        RPA.Logger.info('貼り付けシートのキャンペーンID　   →　'+JudgeCampaignId);
        RPA.Logger.info('キャンペーンID一致、広告作成を開始します');
      }
    } catch {
      break;
    }
  }
  const CampaignName = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[2]/div[3]/div/table/tbody/tr[1]/td[3]');
  await RPA.WebBrowser.mouseClick(CampaignName);
  await RPA.sleep(3000);
  // 右下の広告作成を押下
  const CreateButton = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div/div/div[2]/div[3]/div/div[2]/div[2]/div'}),8000);
  await RPA.WebBrowser.mouseClick(CreateButton);
  await RPA.sleep(3000);
  // クリエイティブIDを入力
  const CreativeId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[1]/div[2]/div/input');
  await RPA.WebBrowser.sendKeys(CreativeId,[WorkData[0][0][20]]);
  await RPA.sleep(5000);
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  try {
    var ApplyFlag = ['false'];
    await RPA.sleep(300);
    const NotCreativeId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[3]/div[2]/div/div'}),1000);
    const NotCreativeIdText = await NotCreativeId.getText();
    if (String(NotCreativeIdText) == '該当する項目が存在しません') {
      RPA.Logger.info(`エラー【${NotCreativeIdText}】` + ' 作業をスキップします');
      const ErrorText = [['エラー', 'クリエイティブID：該当する項目が存在しません']];
      await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AH${Row[0]}`,values:ErrorText});
      Start();
    }
  } catch {
    ApplyFlag[0] = 'true';
    RPA.Logger.info('次の処理に進みます');
  }
  // クリエイティブ選択の「選択」をクリック
  const SelectCreative = await RPA.WebBrowser.findElementByXPath('//*[@id="reactroot"]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]/td[1]/div');
  await RPA.WebBrowser.mouseClick(SelectCreative);
  // 配信期間をクリック
  const DateRange = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[2]/div[2]/div/div/div[1]/div');
  await RPA.WebBrowser.mouseClick(DateRange);
  // 配信期間：開始(Z列)を入力
  const DateRangeStart = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[1]/div[1]/input');
  await DateRangeStart.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(DateRangeStart,[WorkData[0][0][25]]);
  await RPA.sleep(300);
  // 配信期間：終了(AA列)を入力
  const DateRangeEnd = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[2]/div[1]/input');
  await DateRangeEnd.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(DateRangeEnd,[WorkData[0][0][26]]);
  await RPA.sleep(300);
  // 配信期間のOKボタンを押下
  const OKButton = await RPA.WebBrowser.findElementByXPath('/html/body/div[2]/div[3]/div/button[1]');
  await RPA.WebBrowser.mouseClick(OKButton);
  // imp比率(AB列)を入力
  if (WorkData[0][0][27].length < 1) {
    const ImpWeight = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[3]/div[2]/div/input');
    await ImpWeight.clear();
  } else {
    const ImpWeight = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[3]/div[2]/div/input');
    await ImpWeight.clear();
    await RPA.WebBrowser.sendKeys(ImpWeight,[WorkData[0][0][27]]);
  }
  // 予約放送枠ID(D列)を入力
  const AdvertisementSlotId = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[4]/div[2]/div/div[1]/input');
  await AdvertisementSlotId.clear();
  await RPA.sleep(100);
  if (WorkData[0][0][3] == 'なし') {
    ;
  } else {
    await RPA.WebBrowser.sendKeys(AdvertisementSlotId,[WorkData[0][0][3]]);
  }
  // 視聴開始単価を入力
  const Money = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[7]/div[2]/div/input');
  await RPA.WebBrowser.sendKeys(Money,['100']);
  await RPA.sleep(300);
}

async function AdJudgeError2(WorkData, Row) {
  // エラーが起きた場合、作業をスキップしてスタートに戻る
  const error = [['エラー']];
  if (WorkData[0][0][20].length < 1) {
    const ErrorText2 = [['【広告作成】クリエイティブIDが入力されていません']];
    RPA.Logger.info(ErrorText2[0][0]);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AX${Row[0]}:AX${Row[0]}`,values:ErrorText2});
  }
  if (WorkData[0][0][25].length < 1 || WorkData[0][0][26].length < 1 || WorkData[0][0][25].length < 1 && WorkData[0][0][26].length < 1) {
    const ErrorText3 = [['【広告作成】必須項目です']];
    RPA.Logger.info(`エラー【${ErrorText3}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AI${Row[0]}:AI${Row[0]}`,values:ErrorText3});
  }
  if (WorkData[0][0][27].length < 1) {
    const ErrorText4 = [['【広告作成】必須入力です']];
    RPA.Logger.info(`エラー【${ErrorText4}】`);
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AS${Row[0]}:AS${Row[0]}`,values:ErrorText4});
  }
  const NotAdvertisementSlotId = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/div[1]/div/div[5]/div[2]/div[1]/div/form/div[2]/div/div[4]/div[2]/div/div[2]'}),1000);
  const NotAdvertisementSlotIdText = await NotAdvertisementSlotId.getText();
  if (String(NotAdvertisementSlotIdText) == '該当する放送枠が存在しません') {
    RPA.Logger.info(`エラー【${NotAdvertisementSlotIdText}】`);
    const ErrorText5 = [['【広告作成】該当する放送枠が存在しません']];
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:error});
    await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AJ${Row[0]}:AJ${Row[0]}`,values:ErrorText5});
  }
}

async function GetAdvertisementId2(Row) {
  // OKボタンをクリック
  const JudgeRow = await RPA.Google.Spreadsheet.getValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`});
  if (JudgeRow[0][0].indexOf('エラー') == 0) {
    RPA.Logger.info('エラーがあるため、作業をスキップします');
    Start();
  }
  const OKButton2 = await RPA.WebBrowser.findElementByXPath('/html/body/div[1]/div/div[5]/div[2]/footer/div[2]');
  await RPA.WebBrowser.mouseClick(OKButton2);
  await RPA.sleep(5000);
  // 発番した広告IDの最右側の「・・・」をマウスオーバー
  const BalloonMenu = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      xpath: '/html/body/div/div/div[2]/div[3]/div/table/tbody/tr/td[9]/div'
    }),
    5000
  );
  await RPA.WebBrowser.mouseMove(BalloonMenu);
  await RPA.sleep(1000);
  // 「配信を変更する」をクリック
  const ChangeDelivery = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[3]/div/div[4]');
  await RPA.WebBrowser.mouseClick(ChangeDelivery);
  await RPA.sleep(1000);
  const OKButton3 = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[5]/div[2]/footer/div[2]');
  await RPA.WebBrowser.mouseClick(OKButton3);
  await RPA.sleep(8000);
  const GetAdvertisementId = await RPA.WebBrowser.findElementByXPath('/html/body/div/div/div[2]/div[3]/div/table/tbody/tr/td[1]');
  const AdvertisementId = await GetAdvertisementId.getText();
  RPA.Logger.info(AdvertisementId);
  // キャンペーン・広告貼り付けシートのBO列に発番した広告IDを記載
  await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName2}!BO${Row[0] - 1}:BO${Row[0] - 1}`,values:[[AdvertisementId]]});
  // 作業完了を記載
  await RPA.Google.Spreadsheet.setValues({spreadsheetId:`${SSID}`,range:`${SSName}!AG${Row[0]}:AG${Row[0]}`,values:[['完了']]});
}
