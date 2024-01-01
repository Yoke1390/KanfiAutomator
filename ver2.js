function create() {
    // GoogleスプレッドシートのAPIを使ってアクティブなスプレッドシートを取得
    const spread = SpreadsheetApp.getActiveSpreadsheet();
    // スプレッドシート内の"メイン" という名前のシートを取得
    const mainSheet = spread.getSheetByName("メイン");

    // ログデータを記録するセルを取得
    const logdata = mainSheet.getRange("d12");
    // ログデータを記録する関数
    function recordLog(text) {
        logdata.setValue(text); // シートの指定されたセルにログを設定
        console.log(text); // コンソールにもログを出力（デバッグ用）
    }

    // スクリプトの実行開始をログに記録
    recordLog("実行開始");

    // アンケートデータが存在するスプレッドシートのURLを取得し、そのURLを使ってシートを取得
    recordLog("アンケート結果の取得中...");
    const kanfiSheet = SpreadsheetApp.openByUrl(mainSheet.getRange("b3").getValue()).getActiveSheet();

    // 感フィのスライドのURLを取得し、そのURLを使ってプレゼンテーションを開く
    recordLog("スライドの取得中...");
    const kanfiSlide = SlidesApp.openByUrl(mainSheet.getRange("d3").getValue());

    // メンバーの名前を取得。Googleフォームの質問項目から取得。3列目まではそのほかの質問。
    const memberNames = kanfiSheet.getRange(1, 4, 1, kanfiSheet.getLastColumn() - 3).getValues()[0];
    // ニックネームを取得
    const nickNamesSource = kanfiSheet.getRange(2, 2, kanfiSheet.getLastRow() - 1, 2).getValues();
    // ニックネームを辞書（連想配列）に格納
    const nickNames = {};
    nickNamesSource.forEach(function (row) {
        nickNames[row[0]] = row[1];
    });

    // メンバーへのメッセージを取得
    const messageSource = kanfiSheet.getRange(2, 4, kanfiSheet.getLastRow() - 1, kanfiSheet.getLastColumn() - 3).getValues();
    const messages = {};
    for (let i = 0; i < memberNames.length; i++) {
        messages[memberNames[i]] = messageSource.map(function (row) {
            return row[i].trim();
        }).filter(function (message) {
            // 空白の文字列をフィルタリング
            return message !== '';
        });
    }

    // メンバーごとに1枚目と2枚目のスライドの識別子を作成。これをスピーカーノートに記入する。
    const memberNotes = memberNames.flatMap(function (name) {
        return [name + "1枚目。このテキストは変更しないでください。", name + "2枚目。このテキストは変更しないでください。"];
    });

    // スライド内の識別子（スピーカーノート）がない全てのスライドを削除する。
    // 同時に、識別子を持つスライドを配列に格納する。
    recordLog("既存のスライドを確認中...");
    const memberSlides = {}; // メンバーごとのスライドを格納するオブジェクト
    const slidesToRemove = []; // 削除するスライドを格納する配列
    const slideList = kanfiSlide.getSlides(); // 現在のスライドを取得
    for (let i = 0; i < slideList.length; i++) {
        let slide = slideList[i];
        // スライドのスピーカーノートを取得し、不要な改行と空白を削除
        let note = slide.getNotesPage().getSpeakerNotesShape().getText().asString().trim().replace("\n", "");
        // スピーカーノートが空の場合はスライドを削除リストに追加
        if (note.replace("\n", "") == "") {
            slidesToRemove.push(slide);
        }
        // スピーカーノートが識別子リストに含まれている場合は、スライドを保存
        else if (memberNotes.includes(note)) {
            memberSlides[note] = slide;
        }
    }
    // 削除リストに含まれるスライドを実際に削除
    slidesToRemove.forEach(function (slide) {
        slide.remove();
    });

    function initSlide(nameText, noteText) {
        // ニックネームを取得
        let nicknameText;
        if (nickNames[nameText]) nicknameText = nickNames[nameText];
        else nicknameText = nameText; // ニックネームがない場合は本名を表示

        // 現在のスライドの中に指定した識別子（スピーカーノート）があるかどうかを判断
        let slideOfName; // 最終的に返すスライド
        let check = false; // スライドの中にニックネームがあるかどうかを判断する変数
        let isNewSlide = false; // 新しくスライドを追加したかどうかを判断する変数
        // 識別子からスライドを取得できる場合は、そのスライドを用いる
        if (memberSlides[noteText]) {
            slideOfName = memberSlides[noteText];
            // 見つかったスライドの中にニックネームがあるかどうかを判断
            let shapes = slideOfName.getShapes();
            shapes.forEach(function (shape) {
                if (shape.getText().asString().trim().replace("\n", "") == nicknameText) {
                    check = true;
                }
            });
        } else { // スライドが見つからない場合、新しいスライドを追加
            slideOfName = kanfiSlide.appendSlide();
            // スライドの識別子（スピーカーノート）を設定
            slideOfName.getNotesPage().getSpeakerNotesShape().getText().setText(noteText);
            isNewSlide = true;
        }
        if (!check || isNewSlide) {
            // スライドの中身が正しくないまたは新しくスライドを追加した場合、
            // スライドにテキストボックスを挿入し、ニックネームを表示
            let personNameShape = slideOfName.insertShape(SlidesApp.ShapeType.TEXT_BOX, 155, 150, 400, 100);
            let personNameText = personNameShape.getText();
            personNameText.setText(nicknameText);
            personNameText.getTextStyle().setForegroundColor("#ffffff").setFontSize(60);
            personNameText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

            // ニックネームでない本名のみのテキストボックスを削除。ただし、新しくスライドを追加した場合は削除しない
            if (nicknameText != nameText && !isNewSlide) {
                let shapes = slideOfName.getShapes();
                shapes.forEach(function (shape) {
                    if (shape.getText().asString().trim().replace("\n", "") == nameText) {
                        shape.remove();
                    }
                });
            }
        }
        return slideOfName;
    }

    // スプレッドシートからアンケート結果を取得し、スライドに表示
    memberNames.forEach(function (nameText) {
        // 1枚目のスライド ////////////////////////////////////////////////////////////////
        recordLog(nameText + "の1枚目のスライドを処理中...");
        // 1枚目のスライドの識別子（スピーカーノート）を作成
        let note1 = nameText + "1枚目。このテキストは変更しないでください。";
        // 1枚目のスライドを作成。1枚目は初期化するだけでいい
        initSlide(nameText, note1);

        // 2枚目のスライド ////////////////////////////////////////////////////////////////
        recordLog(nameText + "の2枚目のスライドを処理中...");
        // 2枚目のスライドの識別子（スピーカーノート）を作成
        let note2 = nameText + "2枚目。このテキストは変更しないでください。";
        // 2枚目のスライドを初期化。後の処理のため、変数に格納
        let slide2 = initSlide(nameText, note2);

        // 現在スライドにあるすべてのテキストをsetに格納。重複を避けるため。
        let messageSet = new Set();
        let shapes = slide2.getShapes();
        shapes.forEach(function (shape) {
            messageSet.add(shape.getText().asString().trim().replace("\n", ""));
        });

        let sgn = 0; // テキストを挿入する位置を指定する変数
        // メンバーへのメッセージをスライドに表示
        messages[nameText].forEach(function (messageText) {
            if (messageSet.has(messageText.replace("\n", ""))) return; // 重複を避ける
            messageSet.add(messageText.trim().replace("\n", ""));
            // スライドにテキストボックスを挿入し、アンケートの回答を表示
            let shape = slide2.insertShape(SlidesApp.ShapeType.TEXT_BOX, ((sgn - (sgn % 8)) / 8) * 150 + 450, (sgn % 8) * 50, 200, 50);
            let textRNG = shape.getText();
            textRNG.setText(messageText);
            // テキストのスタイルを設定。色をランダムに設定する。
            textRNG.getTextStyle().setForegroundColor(225 - Math.ceil(Math.random() * 4) * 30, 225 - Math.ceil(Math.random() * 4) * 30, 225 - Math.ceil(Math.random() * 4) * 30).setFontSize(20);
            sgn++;
        });
    });
    recordLog("処理完了");
}
