// アンケートの作成関数の引数、指定した日付からdaysの日数分作られる
// year(int)：今年の数値,first_day(string)：アンケートの最初の日付,days(int)：最初の日からの日数,pos_list(list)：どのタイミングに入ってもらうか,holiday(int)：0：全ての日にち,1：土日のみ,2：土日祝のみ
function kaituke_anc(year,first_day,days,pos_list,holiday) {
// フォームの作成、指定フォルダへ保存
  let Form = FormApp.create(`飼付${year}/${first_day}~`).setTitle("飼付アンケート🐎").setDescription("送信ミスや予定変更により回答し直したい場合は、もう一度回答してもらって構いませんがなるべくミスの無いようお願いします。また、自分の名前以外での回答もご遠慮ください。");
//   let id = PropertiesService.getScriptProperties().getProperty('1XcHDN-Z2cBXZEWYUp5WInHZJXFwlG_zM');
//   let formFile = DriveApp.getFileById(Form.getId());
// // 指定したフォルダへ保存
//   DriveApp.getFolderById(id).addFile(formFile);
// // デフォルトで作られていたものの削除
//   DriveApp.getRootFolder().removeFile(formFile);
// 名前を選択するためのプルダウンの作成
  let ListItem = Form.addListItem();
  ListItem.setTitle('名前を選択してください');
  ListItem.setRequired(true);
  ListItem.setChoiceValues(['Aさん','Bさん','Cさん','Dさん','Eさん','Fさん','Gさん','Hさん'])

// グリッドの作成、タイトル設定  
  let GridItem = Form.addGridItem();
  GridItem.setTitle('飼付アンケート');
// 行の設定
  let day_list = [];
// 引数の日付から最初の月と日を変数に
  let monthandday = first_day.split("/")
  let month = monthandday[0];
  let day = monthandday[1];

  // let date = new Date(year,month-1,day);
  // for (let i = 0;i < days;i++){
  //   day_list.push(Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd'));
  //   date.setDate(date.getDate() + 1);
  // }
    

// 必要分の日付のリスト作成
  if (holiday == 0){
    var date = new Date(year,month-1,day);
    for (let i = 0;i < days;i++){
      day_list.push(Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd'));
      date.setDate(date.getDate() + 1);
    }
  }
  else if (holiday == 1){
    var date = new Date(year,month-1,day);
    for (let i = 0;i < days;i++){
      if (date.getDay() == 0 || date.getDay() == 6)
        day_list.push(Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd'));
      date.setDate(date.getDate() + 1);
    }
  }
  
  console.log(day_list);

// 回答すべき飼付の行リストを作成
  let pos_length = pos_list.length;
  let days_length = day_list.length;
  let check_list = [];
  
  for (let i = 0;i < days_length;i++){
    for (let k = 0;k < pos_length;k++){
      check_list.push(day_list[i] + '  ' + pos_list[k]);
    }
  }

  GridItem.setRows(check_list);

// 列の設定  
  GridItem.setColumns(['⭕️','❌']);

// 各行解答することを必須としている
  GridItem.setRequired(true);
}

// 必ずmain関数の方を実行する、上の実行の設定に注意
function main(){
  kaituke_anc(2022,'12/03',23,['朝','夕','夜'],1);
}
