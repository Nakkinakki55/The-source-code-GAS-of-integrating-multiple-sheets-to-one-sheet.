function copyPaste() {
	//シート指定
	var spreadsheet = SpreadsheetApp.getActive();
//↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓change↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
//1つに統合したいシート名を指定する(シート名が違えば名前を変更する必要あり)

	spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DataForIntegrating'), true);
	//↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓change↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓



	//一番先頭のセル
	var rngSheet1 = spreadsheet.getRange("A2");
	rngSheet1.activate();

	// そのセル範囲が空白かどうか判定しログに出力
	if ((rngSheet1.isBlank())) {
		//メッセージボックスを出す
		var result = Browser.msgBox("セルが空白です", Browser.Buttons.OK_CANCEL);
		if (result == "cancel") {
			Logger.log("canceled...")
		}
		//空白なら処理を抜ける
		return;
	}

	//While文　空白になるまで/
	while ((rngSheet1.isBlank()) === false)

	{
		//セルの値取得
		var strValue = rngSheet1.getValue();
		var activeSheetName = SpreadsheetApp.getActiveSpreadsheet();
		//↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓change↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

		//1つに統合したいシート名を指定する
		spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Data1'), true);

		//検索　取得した値を入れる
		var sheet = spreadsheet.getSheetByName('Data1');
		//↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑change↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑-

		var textFinder = sheet.createTextFinder

		var textFinder = sheet.createTextFinder(strValue);
		var ranges = textFinder.findAll();
		var strRange = null;


		for (var i = 0; i < ranges.length; i++) {
			strRange = ranges[i].getA1Notation();
		}


		//もし検索結果があるなら
		if (strRange !== null) {
			//検索したところに飛ぶ
			var rng = spreadsheet.getRange(strRange);
			rng.activate();
			//↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓change↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
			//コピー
			var strValue22222 = rng.offset(0, 1).getValue();
			
			//元のシートに移動
			// 現在アクティブなスプレッドシートを取得
			spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DataForIntegrating'), true); // 現在アクティブなスプレッドシートを取得

			//セルを右に移動
			rngSheet1.offset(0,2).activate();
			//↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑change↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

			///////////////////////////////////////////
			//貼り付け
			spreadsheet.getCurrentCell().setValue(strValue22222);
		} else {

			//↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓change↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
			// 現在アクティブなスプレッドシートを取得
			spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DataForIntegrating'), true); // 現在アクティブなスプレッドシートを取得

			//セルを右に移動
			rngSheet1.offset(0,2).activate();
			//↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑change↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
			//貼り付け
			spreadsheet.getCurrentCell().setValue("");

			spreadsheet.getCurrentCell().setBackground('#ffff00');


		}


		//セルを左に下に移動する
		rngSheet1.offset(1, 0).activate();
		//そのセルを空白判定をする
		var strNextRange = spreadsheet.getCurrentCell().getA1Notation();
		rngSheet1 = spreadsheet.getRange(strNextRange);

	}

	if ((rngSheet1.isBlank())) {
		//メッセージボックスを出す
		var result = Browser.msgBox("Finish！", Browser.Buttons.OK_CANCEL);
		if (result == "cancel") {
			Logger.log("canceled...")
		}

	}
}