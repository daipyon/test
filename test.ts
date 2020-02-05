import RPA from "ts-rpa";

// 読み込みする スプレッドシートID と シート名 の記載
const SSID = process.env.Bansen_ID2_SheetID;
const SSName1 = process.env.Youtube_Delete_Request_SheetName;

RPA.Logger.info(SSID);