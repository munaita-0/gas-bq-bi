function updateHandler() {
    var result = Browser.msgBox("クエリを実行します", Browser.Buttons.OK_CANCEL);
    if (result == "ok"){
      initDashboard().update();
      Browser.msgBox("完了しました", Browser.Buttons.OK_CANCEL);
    }
}
