//Moment.jsライブラリをロード　プロジェクトコード：MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48

/*
■関数名■
　getFormData
■説明■
　フォームから入力されたデータを指定された月・日・項目に入力するマクロ
■引数■
　データ名：var_FormData    型：Var　　説明：フォームから送られてきたデータ
*/
function getFormData(var_FormData) {
  
  //定数
  const GETFORM_YEAR="Year";
  const GETFORM_MONTH="Month";
  //const GETFORM_DAY="Day";
  const GETFORM_DATE="日付";
  const GETFORM_NAME = "ショップ名や買い物内容";
  const GETFORM_TYPE = "項目";
  const GETFORM_PRICE = "金額";
  const SHEETNMCONF = "Config";
  const POS_NOWYEAR = "B1";
  const POS_LASTYEAR = "B2";
  const POS_DATE_ROW = 2;
  const POS_NAME_COL = 1;      //Shop or　Materialのカラム位置
  const POS_TYPE_COL = 2;      //項目のカラム位置
  const TYPENM1 = String("食材費");
  const TYPENM2 = String("日用品");
  
  //変数
  var int_PosDate = 0;
  var int_PosTypeRow = 0;
  var bo_ShopOrMaterials_flg = Boolean(0);
  
  
  //処理
  //var str_GetFormYear = var_FormData.namedValues[GETFORM_YEAR];     //Yearを取得
  //var str_GetFormMonth = var_FormData.namedValues[GETFORM_MONTH];   //Monthを取得
  //var str_GetFormDay = var_FormData.namedValues[GETFORM_DAY];       //Dayを取得
　 　var str_GetFormDate = var_FormData.namedValues[GETFORM_DATE];     //日付を取得
  
  var str_GetFormName = var_FormData.namedValues[GETFORM_NAME];     //ショップ名や買い物内容を取得

  var str_GetFormType = String(var_FormData.namedValues[GETFORM_TYPE]);     //項目を取得

  var str_GetFormPrice = var_FormData.namedValues[GETFORM_PRICE];   //金額を取得

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();          //Container Bound Scriptなので必要ないかもしれないが念のため。
  
  var confsheet = spreadsheet.getSheetByName(SHEETNMCONF);          //Configシート名を指定する
  
  //Year(Now)を取得
  var str_NowYear  = confsheet.getRange(POS_NOWYEAR).getValue();
  
  //Year(Last)を取得
  var str_LastYear = confsheet.getRange(POS_LASTYEAR).getValue();
  
  
  //更新対象シート名を取得
  //var str_SheetNm = getTargetSheetNM(str_GetFormYear,str_GetFormMonth,str_GetFormDay);
  var str_SheetNm = getTargetSheetNM(str_GetFormDate);
  var registSheet = spreadsheet.getSheetByName(str_SheetNm);

  
  //小計のカラム番号を取得する
  var int_PosSmallSum = getPosSmallSum(str_SheetNm,str_NowYear);
  
  //Debug
  //Logger.log(str_SheetNm);

  //指定された日付
  //var str_TargetDate = str_GetFormMonth + "/" + str_GetFormDay;
  //var str_TargetDate = Utilities.formatDate(new Date(str_GetFormDate), ‘Asia/Tokyo’, ‘yyyy/MM/dd’);
  var m = Moment.moment(str_GetFormDate,"YYYYMMDD");
  var str_TargetDate= m.format('M/D');

  
  //2行目C列（３）～AG列(33)まで。i++なので
  for(var i = 3;i <= 33; i++){
  
    var str_SearchDate  = registSheet.getRange(POS_DATE_ROW,i).getValue();
    
    if (str_TargetDate == str_SearchDate) {
      
      int_PosDate = i;
      break;
      
    }
    
  }

  
  //Debug
  //spreadsheet.getSheetByName("Config").getRange("A11").setValue("int_PosDate:" + int_PosDate);
  
  //行を特定するために項目をチェック
  if (str_GetFormType == TYPENM1){
    
    //食材費
    int_PosTypeRow = 5;
    
  } else if(str_GetFormType == TYPENM2){

    //日用品
    int_PosTypeRow = 6;
    
  } else {


    //食材費と日用品以外
    var i = Number(7);

    //debug
    //spreadsheet.getSheetByName("Config").getRange("A12").setValue("それ以外Point1通過: int_PosTypeRow:" + int_PosTypeRow + "  bo_ShopOrMaterials_flg:" + bo_ShopOrMaterials_flg + " Name: " + registSheet.getRange(i, POS_NAME_COL).getValue() + + " Type: " + registSheet.getRange(i, POS_TYPE_COL).getValue());

    while(registSheet.getRange(i, POS_NAME_COL).getValue() != "") {
      
      //Shop or Materialが一致しているか判定
      if (str_GetFormName == registSheet.getRange(i, POS_NAME_COL).getValue() && str_GetFormType == registSheet.getRange(i, POS_TYPE_COL).getValue()){
      
        int_PosTypeRow = i;
        bo_ShopOrMaterials_flg = true;

        //debug
        //spreadsheet.getSheetByName("Config").getRange("A13").setValue("Shop or Materialが一致しているか判定内： int_PosTypeRow:" + int_PosTypeRow + " bo_ShopOrMaterials_flg:" + bo_ShopOrMaterials_flg);

        break;
      }

      i++;
    }
    
    //Shop Or Materialの検索にヒットしたかを判定
    if (bo_ShopOrMaterials_flg == false) {
    
      //debug
      //spreadsheet.getSheetByName("Config").getRange("A14").setValue("bo_ShopOrMaterials_flg==false通過： int_PosTypeRow:" + int_PosTypeRow + "  bo_ShopOrMaterials_flg:" + bo_ShopOrMaterials_flg + ":" + i );

      //i行目に一行を追加する
      registSheet.insertRowAfter(i);
      
      //Shop or Materialに入力
      registSheet.getRange(i, POS_NAME_COL).setValue(str_GetFormName);
      
      //Typeに入力
      registSheet.getRange(i, POS_TYPE_COL).setValue(str_GetFormType);
      
      //6行目の小計のセルからコピー先の行iに計算式をコピーしてくる
      registSheet.getRange(6,int_PosSmallSum).copyTo(registSheet.getRange(i,int_PosSmallSum));  
      
      //追加した行を更新対象行とする
      int_PosTypeRow = i;
      
    }

    //debug
    //spreadsheet.getSheetByName("Config").getRange("A15").setValue("elseの最後!! int_PosTypeRow: " + int_PosTypeRow + "  bo_ShopOrMaterials_flg:" + bo_ShopOrMaterials_flg );

  }
  //debug
  //spreadsheet.getSheetByName("Config").getRange("A15").setValue("対象セルの位置!! int_PosTypeRow: " + int_PosTypeRow + "  int_PosDate:　" + int_PosDate );
 
  //更新対象セルに値を入力する
  registSheet.getRange(int_PosTypeRow,int_PosDate).setValue(str_GetFormPrice);
  
  //初期化
  i = 0;
}


/*
■関数名■
　getTargetSheetNM
■説明■
　渡された月日からシート名を返す
■引数■
1.  データ名：str_Year    型：String　　 説明：年
2.  データ名：str_Month    型：String　　説明：月
3.  データ名：str_Day      型：String　　説明：日
■戻り値■
　型：String　説明：月日に対応したシート名
*/
function getTargetSheetNM(str_Date) {

  //テスト用
  //var str_Year = 2019
  //var str_Month = 12
  //var str_Date = 24

  //定数
  const SHEETNMCONF = "Config";
  const POS_NOWYEAR = "B1";
  const POS_LASTYEAR = "B2";

  //変数

  
  //処理
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();           //Container Bound Scriptなので必要ないかもしれないが念のため。
  var confsheet = spreadsheet.getSheetByName(SHEETNMCONF);        //Configシート名を指定する

  //test
  //spreadsheet.getSheetByName("Config").getRange("A19").setValue(str_Date);
  
  //Year(Now)を取得
  var str_NowYear  = confsheet.getRange(POS_NOWYEAR).getValue();
  
  //Year(Last)を取得
  var str_LastYear = confsheet.getRange(POS_LASTYEAR).getValue();


  
  //引数に渡された日付を取得
  //var str_TargetDate = new Date(str_Year + "/" + str_Month + "/" + str_Date); 
  var str_TargetDate = new Date(str_Date); 

  //12/25~1/24まではJunuaryシート
  var startdate  = new Date(str_LastYear + "/12/25");
  var enddate  = new Date(str_NowYear + "/1/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
    var str_SheetNm = "January";
    return str_SheetNm;
  } 
  
  //1/25~2/24まではFebruaryシート
  var startdate  = new Date(str_NowYear + "/1/25");
  var enddate  = new Date(str_NowYear + "/2/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
    var str_SheetNm = "February";
    return str_SheetNm;
  } 

  //2/25~3/24まではMarchシート
  var startdate  = new Date(str_NowYear + "/2/25");
  var enddate  = new Date(str_NowYear + "/3/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
    var str_SheetNm = "March";
    return str_SheetNm;
  } 

  //3/25~4/24まではAprilシート
  var startdate  = new Date(str_NowYear + "/3/25");
  var enddate  = new Date(str_NowYear + "/4/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
     var str_SheetNm = "April";
    return str_SheetNm;
  } 

  //4/25~5/24まではMayシート
  var startdate  = new Date(str_NowYear + "/4/25");
  var enddate  = new Date(str_NowYear + "/5/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
      var str_SheetNm = "May";
    return str_SheetNm;
  } 

  //5/25~6/24まではJuneシート
  var startdate  = new Date(str_NowYear + "/5/25");
  var enddate  = new Date(str_NowYear + "/6/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
    var str_SheetNm = "June";
    return str_SheetNm;
  } 
      
  //6/25~7/24まではJulyシート
  var startdate  = new Date(str_NowYear + "/6/25");
  var enddate  = new Date(str_NowYear + "/7/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
    var str_SheetNm = "July";
    return str_SheetNm;
  } 
      
  //7/25~8/24まではAugustシート
  var startdate  = new Date(str_NowYear + "/7/25");
  var enddate  = new Date(str_NowYear + "/8/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
    var str_SheetNm = "August";
    return str_SheetNm;
  } 
      
  //8/25~9/24まではSeptemberシート
  var startdate  = new Date(str_NowYear + "/8/25");
  var enddate  = new Date(str_NowYear + "/9/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
    var str_SheetNm = "September";
    return str_SheetNm;
  } 
      
  //9/25~10/24まではOctorberシート
  var startdate  = new Date(str_NowYear + "/9/25");
  var enddate  = new Date(str_NowYear + "/10/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
    var str_SheetNm = "Octorber";
    return str_SheetNm;
  } 
      
  //10/25~11/24まではOctorberシート
  var startdate  = new Date(str_NowYear + "/10/25");
  var enddate  = new Date(str_NowYear + "/11/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
    var str_SheetNm = "November";
    return str_SheetNm;
  } 
      
      
  //11/25~12/24まではDecemberシート
  var startdate  = new Date(str_NowYear + "/11/25");
  var enddate  = new Date(str_NowYear + "/12/24" );
  if(str_TargetDate >= startdate && str_TargetDate <= enddate) { 
    var str_SheetNm = "December";
    return str_SheetNm;
  } 
  
  //return str_SheetNm;  
  //Logger.log(str_SheetNm)

}


/*
■関数名■
　getPosSmallSum
■説明■
　引き渡されたシート名を元にカラム番号を返す
■引数■
1.  データ名：str_SheetNm    型：String　　 説明：シート名
2.  データ名:var_NowYear    型：Number 　説明：年
■戻り値■
　型：Number　意味：小計のカラム番号
*/
function getPosSmallSum(str_SheetNm,var_NowYear) {
  
  var int_UruuToshi;  //閏年除算用変数
  
  if (str_SheetNm == 'Junuary'){
    
    //AHカラム
    return 34;
    
  } else if (str_SheetNm == 'February'){
    
    //AHカラム
    return 34;

  } else if (str_SheetNm == 'March'){
    
    //指定された年が閏年かどうかを判定する
    int_UruuToshi = var_NowYear % 4
    
    if (var_NowYear > 0 ){ 
        
        //AFカラム
        return 32;

    } else if(var_NowYear == 0 ){
    
        //AEカラム
        return 31;
    
    }

  } else if (str_SheetNm == 'April'){

    //AHカラム
    return 34;

  } else if (str_SheetNm == 'May'){

    //AGカラム
    return 33;

  } else if (str_SheetNm == 'June'){

    //AHカラム
    return 34;

  } else if (str_SheetNm == 'July'){

    //AGカラム
    return 33;

  } else if (str_SheetNm == 'August'){

    //AHカラム
    return 34;

  } else if (str_SheetNm == 'September'){

    //AHカラム
    return 34;

  } else if (str_SheetNm == 'Octorber'){

    //AGカラム
    return 33;

  } else if (str_SheetNm == 'November'){

    //AHカラム
    return 34;

  } else if (str_SheetNm == 'December'){

    //AGカラム
    return 33;


  }
  
}

