class Member {
    constructor(name) {
        this.name = name;
        this.note1 = name + "1枚目。このテキストは変更しないでください。"; // 1枚目のスライドの識別子（スピーカーノート）
        this.note2 = name + "2枚目。このテキストは変更しないでください。"; // 2枚目のスライドの識別子（スピーカーノート）
        this.nickname = name; // ニックネームの初期値を本名にする
        this.messages = [];
        this.slide1 = null;
        this.slide2 = null;
        this.slides = [this.slide1, this.slide2];
    }
}

function main() {
    // 感フィ君のスプレッドシートを取得。
    const main_sheet = connectMainSheet();
    // D12のセルにログを記録する関数を作成
    const log = createLogger(main_sheet.getRange("d12"));
    log("実行開始");

    // データの処理 ///////////////////////////////////////////////////
    log("アンケート結果の取得中...");
    const { member_name_source, nickname_source, message_source } = fetchKanfiData(main_sheet);
    log("メンバーの一覧を作成中...");
    const members = makeMembersMap(member_name_source);
    log("読んでほしい名前を読み込み中...");
    setNickNames(nickname_source, members);
    log("メッセージを割り当て中...");
    setMessages(message_source, members);

    // スライドの処理 /////////////////////////////////////////////////
    log("感フィを作成するスライドと接続中...");
    const kanfi_slide = connectKanfiSlide(main_sheet);
    log("既存のスライドを割り当て中...");
    assignExsistingSlidestoMembers(members, kanfi_slide);
    for (const member in members) {
        log(member.name + "の1枚目のスライドを処理中...");
        firstSlide(member, kanfi_slide);
        log(member.name + "の2枚目のスライドを処理中...");
        secondSlide(member, kanfi_slide);
    };
    //  TODO: スライドの順番を変更
    // sortSlides(kanfi_slide);
}

// データの処理 //////////////////////////////////////////////////////////////////////////////

function connectMainSheet() {
    // GoogleスプレッドシートのAPIを使ってアクティブなスプレッドシートを取得
    const spread = SpreadsheetApp.getActiveSpreadsheet();
    // スプレッドシート内の"メイン" という名前のシートを取得
    const main_sheet = spread.getSheetByName("メイン");
    return main_sheet;
}

function fetchKanfiData(mainSheet) {
    // アンケートデータが存在するスプレッドシートのURLを取得し、そのURLを使ってシートを取得
    const form_url = mainSheet.getRange("b3").getValue();
    // 感フィのスライドのURLを取得し、そのURLを使ってプレゼンテーションを開く
    const kanfi_form_sheet = SpreadsheetApp.openByUrl(form_url).getActiveSheet();

    // let xxx_source... : データの取得
    // xxx_source=xxx_source.map... : データの整形(空白の削除)

    let member_name_source = kanfi_form_sheet.getRange(1, 4, 1, kanfi_form_sheet.getLastColumn() - 3).getValues()[0];
    member_name_source = member_name_source.map(
        name => deleteSpace(name)
    );

    let nickname_source = kanfi_form_sheet.getRange(2, 2, kanfi_form_sheet.getLastRow() - 1, 2).getValues();
    nickname_source = nickname_source.map(
        row => row.map(text => deleteSpace(text))
    );

    let message_source = kanfi_form_sheet.getRange(2, 4, kanfi_form_sheet.getLastRow() - 1, kanfi_form_sheet.getLastColumn() - 3).getValues();
    message_source = message_source.map(
        row => row.map(message => deleteSpace(message))
    );

    return { member_name_source, nickname_source, message_source };
}

function makeMembersMap(member_name_source) {
    const members = new Map();
    for (const name of member_name_source) {
        members.set(name, new Member(name));
    };
    return members;
}

function setNickNames(nickname_source, members) {
    for (const row of nickname_source) {
        const name = row[0];
        const target_member = members.get(name);

        const nickname = row[1];
        target_member.nickname = nickname;
    };
    // 正しいニックネームが設定されているか確認
    // for (const name in members) {
    //     console.log(members[name].nickname); // テスト用なので、コンソールに出力すればOK
    // };
}

function setMessages(message_source, members) {
    let i = 0;
    for (const member of members.values()) {
        // membersの並び順とmessage_sourceの並び順が一致していることを前提としている
        member.messages = message_source.map(function (row) {
            return deleteSpace(row[i]);
        }).filter(function (message) {
            return message !== "";
        });
        i++;
    }
    // 正しいメッセージが設定されているか確認するためのコード。
    // for (const name in members) {
    //     console.log(members[name].messages); // テスト用なので、コンソールに出力すればOK
    // };
}

// スライドの処理 //////////////////////////////////////////////////////////////////////////////

function connectKanfiSlide(mainSheet) {
    // D3のセルに記載されているURLを使ってスライドを開く
    return SlidesApp.openByUrl(mainSheet.getRange("d3").getValue());
}

function assignExsistingSlidestoMembers(members, kanfi_slide) {
    // 各メンバーのスライドの識別子を格納する集合を作成
    const set_of_notes = new Set();
    for (const name in members) {
        set_of_notes.add(members[name].note1);
        set_of_notes.add(members[name].note2);
    }

    const exsisting_slides = kanfi_slide.getSlides();
    const slides_to_remove = []; //  todo：削除リストが必要かどうかテスト
    for (const slide of exsisting_slides) {
        // fixme: noteがundefinedになることがある
        const note = deleteSpace(slide.getNotesPage().getSpeakerNotesShape().getText().asString()); // スライドのスピーカーノートを取得
        if (set_of_notes.includes(note)) {
            // スピーカーノートが識別子リストに含まれている場合は、スライドを保存
            setSlideFromNote(note, slide, members); // todo: 動作テスト
        } else {
            // そうでない場合はスライドを削除リストに追加
            log("スピーカーノート「" + note + "」のスライドを削除");
            slides_to_remove.push(slide);
        }
    }
    for (const slide of slides_to_remove) {
        // 削除リストに含まれるスライドを実際に削除
        // slide.remove();
    };
}

function setSlideFromNote(note, slide, members) {
    const name = note.split("枚目")[0].slice(0, -1);
    const slide_index = note.split("枚目")[0].slice(-1) - 1;
    members[name].slides[slide_index] = slide;
}

function firstSlide(member, kanfi_slide) {
    if (member.slide1 == null) {
        member.slide1 = kanfi_slide.appendSlide();
        member.slide1.getNotesPage().getSpeakerNotesShape().getText().setText(member.note1);
    }
    putNicknameOnSlide(member.name, member.nickname, member.slide1);
}

function secondSlide(member, kanfi_slide) {
    // hack: 1枚目のスライドを作成するコードをコピーしている。slide1とslide2を配列にするといいかもしれない。
    if (member.slide2 == null) {
        member.slide2 = kanfi_slide.appendSlide();
        member.slide2.getNotesPage().getSpeakerNotesShape().getText().setText(member.note2);
    }
    putNicknameOnSlide(member.name, member.nickname, member.slide2);
    putMessagesOnSlide(member.messages, member.slide2);
}

function putNicknameOnSlide(name, nickname, slide) {
    const shapes_list = slide.getShapes();
    for (const shape of shapes_list) {
        if (deleteSpace(shape.getText().asString()) == nickname) {
            // ニックネームがある場合はそのまま返す
            return slide;
        }
    };

    // ニックネームがない場合はニックネームを追加
    const nickname_textbox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 155, 150, 400, 100);
    const textbox_content = nickname_textbox.getText();
    textbox_content.setText(nickname);
    textbox_content.getTextStyle().setForegroundColor("#ffffff").setFontSize(60);
    textbox_content.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // ニックネームでない本名のみのテキストボックスを削除。
    if (nickname != name) {
        for (const shape of shapes_list) {
            if (deleteSpace(shape.getText().asString()) == name) {
                shape.remove();
            }
        };
    }
    return slide;
}

function putMessagesOnSlide(message_list, slide) {
    // 重複を避けるため、現在スライドにあるすべてのテキストをsetに格納。
    const exsisting_messages = new Set(slide.getShapes().map(shape => deleteSpace(shape.getText().asString())));

    let i = 0; // テキストを挿入する位置を指定する変数
    for (message of message_list) {
        if (exsisting_messages.has(message)) continue; // 重複を避ける

        exsisting_messages.add(message);
        putNewMessage(message, slide, i);
        i++;
    };
}

function putNewMessage(message, slide, i) {
    // 8メッセージで1列になるように配置
    const x_coordinate = ((i - (i % 8)) / 8) * 150 + 450;
    const y_coordinate = (i % 8) * 50;
    const textbox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x_coordinate, y_coordinate);

    const textbox_content = textbox.getText();
    textbox_content.setText(message);

    const style = textbox_content.getTextStyle();
    style.setForegroundColor(getRandomRGB());
    style.setFontFamily("HiraginoSans-W3")
    style.setFontSize(20);
}

// その他の処理 //////////////////////////////////////////////////////////////////////////////

function createLogger(loggerCell) {
    // ログデータを記録する関数を作成する関数
    return function (text) {
        loggerCell.setValue(text); // シートの指定されたセルにログを設定
        console.log(text); // コンソールにもログを出力（デバッグ用）
    };
}

function deleteSpace(text) {
    // 正規表現を用いて空白を削除
    return text.replace(/[\s　]+/g, "");
}

function getRandomRGB() {
    // ランダムな色相
    const hue = Math.floor(Math.random() * 360);
    // 最大彩度
    const saturation = 100;
    // 最大明度
    const brightness = 100;

    // RGBに変換
    const rgb = hsvToRgb(hue, saturation, brightness);

    // カラーコードを生成
    const rgb_code = `rgb(${rgb.r}, ${rgb.g}, ${rgb.b})`;

    return rgb_code;
}

// HSBからRGBに変換
function hsvToRgb(hue, saturation, brightness) {
    const h = hue / 360;
    const s = saturation / 100;
    const v = brightness / 100;

    const c = v * s;
    const hh = h * 6;
    const x = c * (1 - Math.abs(hh % 2 - 1));
    const m = v - c;

    let r, g, b;

    if (hue < 60) {
        r = c;
        g = x;
        b = 0;
    } else if (hue < 120) {
        r = x;
        g = c;
        b = 0;
    } else if (hue < 180) {
        r = 0;
        g = c;
        b = x;
    } else if (hue < 240) {
        r = 0;
        g = x;
        b = c;
    } else if (hue < 300) {
        r = x;
        g = 0;
        b = c;
    } else {
        r = c;
        g = 0;
        b = x;
    }

    const rgb = {
        r: Math.round((r + m) * 255),
        g: Math.round((g + m) * 255),
        b: Math.round((b + m) * 255),
    };

    return rgb;
}
