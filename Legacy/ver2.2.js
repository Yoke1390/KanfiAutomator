function main() {
    // 感フィ君のスプレッドシートを取得。対象のフォームのデータやスライドのURLを取得するほか、ログを記録するためのセルも取得
    const mainSheet = initSheet();
    // D12のセルにログを記録する関数を作成
    const recordLog = createLogger(mainSheet.getRange("d12"));
    recordLog("実行開始");

    // 感フィフォームのデータのあるスプレッドシートと、感フィを作成するスライドを取得
    const { kanfiSheet, kanfiSlide } = getKanfiData(mainSheet, recordLog);
    // 感フィのフォームからメンバーの名前、呼んで欲しい名前、メッセージを取得
    const { memberNames, nickNames, messages } = getMemberData(kanfiSheet);
    // スライドの識別子として使うスピーカーノートを作成
    const memberNotes = memberNames.flatMap(function (name) {
        return [name + "1枚目。このテキストは変更しないでください。", name + "2枚目。このテキストは変更しないでください。"];
    });

    const memberSlides = scanExistingSlides(kanfiSlide, memberNotes, recordLog);

    memberNames.forEach(function (nameText) {
        recordLog(nameText + "の1枚目のスライドを処理中...");
        let note1 = nameText + "1枚目。このテキストは変更しないでください。"; // 1枚目のスライドの識別子（スピーカーノート）を作成
        initSlide(nameText, note1, nickNames, kanfiSlide, memberSlides); // 1枚目のスライドを作成。1枚目は初期化するだけでいい

        recordLog(nameText + "の2枚目のスライドを処理中...");
        let note2 = nameText + "2枚目。このテキストは変更しないでください。"; // 2枚目のスライドの識別子（スピーカーノート）を作成
        let slide2 = initSlide(nameText, note2, nickNames, kanfiSlide, memberSlides); // 2枚目のスライドを初期化。後の処理のため、変数に格納
        showMessage(slide2, messages, nameText); // 2枚目のスライドにメッセージを表示
    });

    recordLog("処理完了");
}

// データの処理 //////////////////////////////////////////////////////////////////////////////

function initSheet() {
    // GoogleスプレッドシートのAPIを使ってアクティブなスプレッドシートを取得
    const spread = SpreadsheetApp.getActiveSpreadsheet();
    // スプレッドシート内の"メイン" という名前のシートを取得
    const mainSheet = spread.getSheetByName("メイン");
    return mainSheet;
}

function getKanfiData(mainSheet, recordLog) {
    // アンケートデータが存在するスプレッドシートのURLを取得し、そのURLを使ってシートを取得
    recordLog("アンケート結果の取得中...");
    // 感フィのスライドのURLを取得し、そのURLを使ってプレゼンテーションを開く
    const kanfiSheet = SpreadsheetApp.openByUrl(mainSheet.getRange("b3").getValue()).getActiveSheet();
    recordLog("スライドの取得中...");
    const kanfiSlide = SlidesApp.openByUrl(mainSheet.getRange("d3").getValue());
    return { kanfiSheet, kanfiSlide };
}

function getMemberData(kanfiSheet) {
    // メンバーの名前を取得する。Googleフォームの質問項目から取得。3列目まではそのほかの質問。
    const memberNames = kanfiSheet.getRange(1, 4, 1, kanfiSheet.getLastColumn() - 3).getValues()[0];
    // メンバー名から空白を削除
    for (let i = 0; i < memberNames.length; i++) {
        memberNames[i] = deleteSpace(memberNames[i]);
    }

    const nickNames = getNickNames(kanfiSheet); // 呼んで欲しい名前を取得
    const messages = getMessages(kanfiSheet, memberNames); // メンバーへのメッセージを取得

    return { memberNames, nickNames, messages };
}

function getNickNames(kanfiSheet) {
    // フォームへの回答から呼んで欲しい名前を辞書（連想配列）に格納
    const nickNamesSource = kanfiSheet.getRange(2, 2, kanfiSheet.getLastRow() - 1, 2).getValues();
    const nickNames = {};
    nickNamesSource.forEach(function (row) {
        const name = deleteSpace(row[0]); // 回答に記載されていた本名から、正規表現を用いて空白を削除
        nickNames[name] = row[1]; // 辞書の本名のキーに対して、ニックネームを格納
    });
    return nickNames;
}

function getMessages(kanfiSheet, memberNames) {
    // フォームへの回答からメンバーへのメッセージを取得
    const messageSource = kanfiSheet.getRange(2, 4, kanfiSheet.getLastRow() - 1, kanfiSheet.getLastColumn() - 3).getValues();
    const messages = {}; // メンバー名をキーにして、メッセージの配列を格納
    for (let i = 0; i < memberNames.length; i++) {
        messages[memberNames[i]] = messageSource.map(function (row) {
            return deleteSpace(row[i]);
        }).filter(function (message) {
            return message !== '';
        });
    }
    return messages;
}

// スライドの処理 //////////////////////////////////////////////////////////////////////////////

function scanexistingslides(kanfislide, memberNotes, recordlog) {
    recordlog("既存のスライドを確認中...");
    const memberSlides = {};
    const slidesToRemove = [];
    const slidelist = kanfislide.getslides();
    for (let i = 0; i < slidelist.length; i++) {
        let slide = slidelist[i];
        const note = deleteSpace(slide.getnotespage().getspeakernotesshape().gettext().asstring()); // スライドのスピーカーノートを取得
        if (memberNotes.includes(note)) {
            // スピーカーノートが識別子リストに含まれている場合は、スライドを保存
            memberSlides[note] = slide;
        } else {
            // そうでない場合はスライドを削除リストに追加
            recordLog("スピーカーノート「" + note + "」のスライドを削除")
            slidesToRemove.push(slide);
        }
    }
    slidesToRemove.forEach(function (slide) {
        // 削除リストに含まれるスライドを実際に削除
        slide.remove();
    });
    return memberSlides;
}

function initSlide(nameText, noteText, nickNames, kanfiSlide, memberSlides) {
    // ニックネームを取得
    let nicknameText;
    if (nickNames[nameText]) nicknameText = nickNames[nameText];
    else nicknameText = nameText; // ニックネームがない場合は本名を表示

    // 現在のスライドの中に指定した識別子（スピーカーノート）があるかどうかを判断
    let slide;
    if (memberSlides[noteText]) {
        // スライドがある場合はそのスライドを確認
        slide = checkSlide(memberSlides[noteText], nameText, nicknameText);
    } else {
        // スライドがない場合は新規作成
        slide = createSlide(nicknameText, noteText, kanfiSlide);
    }
    return slide;
}

function checkSlide(slide, nameText, nicknameText) {
    let includeNickName = false;
    let shapes = slide.getShapes();
    shapes.forEach(function (shape) {
        if (deleteSpace(shape.getText().asString()) == nicknameText) {
            // ニックネームがある場合はそのまま返す
            includeNickName = true;
        }
    });
    if (!includeNickName) {
        // ニックネームがない場合はニックネームを追加
        let personNameShape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 155, 150, 400, 100);
        let personNameText = personNameShape.getText();
        personNameText.setText(nicknameText);
        personNameText.getTextStyle().setForegroundColor("#ffffff").setFontSize(60);
        personNameText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        // ニックネームでない本名のみのテキストボックスを削除。
        if (nicknameText != nameText) {
            let shapes = slide.getShapes();
            shapes.forEach(function (shape) {
                if (deleteSpace(shape.getText().asString()) == nameText) {
                    shape.remove();
                }
            });
        }
    }
    return slide;
}

function createSlide(nicknameText, noteText, kanfiSlide) {
    let newSlide = kanfiSlide.appendSlide();
    newSlide.getNotesPage().getSpeakerNotesShape().getText().setText(noteText);
    // スライドにテキストボックスを挿入し、ニックネームを表示
    let personNameShape = newSlide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 155, 150, 400, 100);
    let personNameText = personNameShape.getText();
    personNameText.setText(nicknameText);
    personNameText.getTextStyle().setForegroundColor("#ffffff").setFontSize(60);
    personNameText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    return newSlide;
}

function showMessage(slide, messages, nameText) {
    // 現在スライドにあるすべてのテキストをsetに格納。重複を避けるため。
    let messageSet = new Set();
    let shapes = slide.getShapes();
    shapes.forEach(function (shape) {
        messageSet.add(deleteSpace(shape.getText().asString()));
    });

    let sgn = 0; // テキストを挿入する位置を指定する変数
    // メンバーへのメッセージをスライドに表示
    messages[nameText].forEach(function (messageText) {
        if (messageSet.has(messageText)) return; // 重複を避ける
        messageSet.add(messageText);
        // スライドにテキストボックスを挿入し、アンケートの回答を表示
        let shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, ((sgn - (sgn % 8)) / 8) * 150 + 450, (sgn % 8) * 50, 200, 50);
        let textRNG = shape.getText();
        textRNG.setText(messageText);
        // テキストのスタイルを設定。色をランダムに設定する。
        textRNG.getTextStyle().setForegroundColor(225 - Math.ceil(Math.random() * 5) * 30, 225 - Math.ceil(Math.random() * 4) * 30, 225 - Math.ceil(Math.random() * 4) * 30).setFontSize(20);
        sgn++;
    });
}

// その他の処理 //////////////////////////////////////////////////////////////////////////////
// ログデータを記録する関数を作成する関数
function createLogger(logdata) {
    return function (text) {
        logdata.setValue(text); // シートの指定されたセルにログを設定
        console.log(text); // コンソールにもログを出力（デバッグ用）
    };
}

function deleteSpace(text) {
    return text.replace(/[\s　]+/g, "");
}
