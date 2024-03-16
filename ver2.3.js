class Member {
    constructor(name) {
        this.name = name;
        this.nickname = name; // ニックネームの初期値を本名にする
        this.messages = [];
        this.slides = [null, null];

        const note1 = name + "1枚目。このテキストは変更しないでください。"; // 1枚目のスライドの識別子（スピーカーノート）
        const note2 = name + "2枚目。このテキストは変更しないでください。"; // 2枚目のスライドの識別子（スピーカーノート）
        this.notes = [note1, note2];
    }
}

function main() {
    // 感フィ君のスプレッドシートを取得。
    const main_sheet = connectMainSheet();
    // D12のセルにログを記録する関数を作成
    log = createLogger(main_sheet.getRange("d12"));
    log("実行開始");

    // データの処理 ///////////////////////////////////////////////////
    log("アンケート結果の取得中...");
    const { member_name_source, nickname_source, message_source } = fetchKanfiData(main_sheet);

    log("メンバーの一覧を作成中...");
    const members = makeMembersMap(member_name_source);
    log("呼んでほしい名前を割り当て中...");
    setNicknames(nickname_source, members);
    log("メッセージを割り当て中...");
    setMessages(message_source, members);

    // スライドの処理 /////////////////////////////////////////////////
    log("感フィを作成するスライドと接続中...");
    const kanfi_slide = connectKanfiSlide(main_sheet);
    log("既存のスライドを割り当て中...");
    assignExsistingSlidesToMembers(members, kanfi_slide);

    for (const member of members.values()) {
        log(member.name + "の1枚目のスライドを処理中...");
        firstSlide(member, kanfi_slide);
        log(member.name + "の2枚目のスライドを処理中...");
        secondSlide(member, kanfi_slide);
    };

    log("処理完了");
}

// データの処理 //////////////////////////////////////////////////////////////////////////////

function connectMainSheet() {
    const active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // スプレッドシート内の"メイン" という名前のシートを取得
    const main_sheet = active_spreadsheet.getSheetByName("メイン");
    return main_sheet;
}

function fetchKanfiData(mainSheet) {
    // アンケートデータが存在するスプレッドシートのURLを取得し、そのURLを使ってシートを取得
    const form_url = mainSheet.getRange("b3").getValue();
    const kanfi_form_sheet = SpreadsheetApp.openByUrl(form_url).getActiveSheet();

    // let xxx_source... : データの取得
    // xxx_source=xxx_source.map... : データの整形(空白の削除)。
    //      mapメソッドは配列の各要素に対して関数を実行し、その結果を新しい配列として返す。
    //      =>はアロー関数。引数を受け取り、その引数を元にして何かを返す関数を簡潔に書くための構文。

    // 1行目の4列目から最終列まで（フォームの質問に書かれているメンバーの名前）
    let member_name_source = kanfi_form_sheet.getRange(1, 4, 1, kanfi_form_sheet.getLastColumn() - 3).getValues()[0];
    member_name_source = member_name_source.map(
        name => deleteSpace(name)
    );

    // フォームへの回答の2列目と3列目から、名前と読んで欲しい名前のペアを取得
    let nickname_source = kanfi_form_sheet.getRange(2, 2, kanfi_form_sheet.getLastRow() - 1, 2).getValues();
    nickname_source = nickname_source.map(
        row => row.map(text => deleteSpace(text))
    );

    // 2行目以降、4列目以降から、メンバーへのメッセージを取得
    let message_source = kanfi_form_sheet.getRange(2, 4, kanfi_form_sheet.getLastRow() - 1, kanfi_form_sheet.getLastColumn() - 3).getValues();
    message_source = message_source.map(
        row => row.map(message => deleteSpace(message))
    );

    return { member_name_source, nickname_source, message_source };
}

function makeMembersMap(member_name_source) {
    const members = new Map();
    for (const name of member_name_source) {
        members.set(name, new Member(name)); // Mapオブジェクトの操作はset/getで行うことに注意
    };
    return members;
}

function setNicknames(nickname_source, members) {
    for (const row of nickname_source) {
        const name = row[0];
        const target_member = members.get(name); // Mapオブジェクトの操作はset/getで行うことに注意
        const nickname = row[1];
        target_member.nickname = nickname;
    };
    // 正しいニックネームが設定されているか確認するためのコード。
    // for (const name in members) {
    //     console.log(members[name].nickname); // テスト用なので、コンソールに出力すればOK
    // };
}

function setMessages(message_source, members) {
    // membersの追加順とmessage_sourceの並び順が一致していることを前提としている。
    // 高々数十人のメンバーなので、計算量は気にしないでいい。
    let i = 0;
    for (const member of members.values()) {
        member.messages = message_source.map(
            row => deleteSpace(row[i])
        ).filter(
            message => message !== '' // 空でないメッセージのみを抽出
        );
        i++;
    }
    // 正しいメッセージが設定されているか確認するためのコード。
    // for (const name in members) {
    //     console.log(members[name].messages); // テスト用なので、コンソールに出力すればOK
    // };
}

// スライドの処理 //////////////////////////////////////////////////////////////////////////////

function connectKanfiSlide(mainSheet) {
    // D3のセルに感フィのスライドのURLが記載されている。
    const kanfi_url = mainSheet.getRange("d3").getValue();
    if (deleteSpace(kanfi_url) == "") {
        // URLが空の場合は新しいスライドを作成
        log("感フィのスライドを新規作成");
        const new_slide = SlidesApp.create("感フィ");

        const new_slide_url = new_slide.getUrl();
        mainSheet.getRange("d3").setValue(new_slide_url);

        return new_slide;
    }
    return SlidesApp.openByUrl(kanfi_url);
}

function assignExsistingSlidesToMembers(members, kanfi_slide) {
    // 各メンバーのスライドの識別子を格納する集合を作成
    const set_of_notes = new Set();
    for (const member of members.values()) {
        for (const note of member.notes) {
            set_of_notes.add(note)
        }
    }

    const exsisting_slides = kanfi_slide.getSlides();
    for (const slide of exsisting_slides) {
        const note = deleteSpace(slide.getNotesPage().getSpeakerNotesShape().getText().asString()); // スライドのスピーカーノートを取得
        if (set_of_notes.has(note)) {
            // スピーカーノートが識別子リストに含まれている場合は、スライドを保存
            setSlideFromNote(note, slide, members);
        } else {
            log("スピーカーノート「" + note + "」のスライドを削除");
            slide.remove(); // 削除したくない場合は、この2行をコメントアウト
        }
    }
}

function setSlideFromNote(note, slide, members) {
    log("スピーカーノート「" + note + "」のスライドをセット");
    const slide_info = note.split("枚目")[0]
    const name = slide_info.slice(0, -1);
    const slide_index = slide_info.slice(-1) - 1;

    const target_member = members.get(name);
    target_member.slides[slide_index] = slide;
}

function firstSlide(member, kanfi_slide) {
    initSlide(member, kanfi_slide, 0)
}

function secondSlide(member, kanfi_slide) {
    initSlide(member, kanfi_slide, 1)
    putMessagesOnSlide(member.messages, member.slides[1]);
}

function initSlide(member, kanfi_slide, slide_index) {
    let target_slide = member.slides[slide_index];
    if (target_slide == null) {
        target_slide = kanfi_slide.appendSlide();
        target_slide.getNotesPage().getSpeakerNotesShape().getText().setText(member.notes[slide_index]); // スピーカーノートを設定
        target_slide.getBackground().setSolidFill("#000000"); // スライドの背景色を黒に設定
        member.slides[slide_index] = target_slide;
    }
    putNicknameOnSlide(member.name, member.nickname, target_slide);
    return target_slide;
}

function putNicknameOnSlide(name, nickname, slide) {
    // ニックネームがある場合はそのまま返す
    const textbox_list = slide.getShapes();
    for (const textbox of textbox_list) {
        if (deleteSpace(textbox.getText().asString()) == nickname) {
            return;
        }
    };

    // ニックネームがない場合はニックネームを追加
    const nickname_textbox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 155, 150, 400, 100);
    const textbox_content = nickname_textbox.getText();
    textbox_content.setText(nickname);
    textbox_content.getTextStyle().setForegroundColor("#ffffff").setFontSize(60);
    textbox_content.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // このメンバーがフォームに回答する前に作成されたスライドには、初期値の本名のテキストボックスが存在する。
    // ニックネームでない本名のみのテキストボックスを削除。
    if (nickname != name) {
        for (const shape of textbox_list) {
            if (deleteSpace(shape.getText().asString()) == name) {
                shape.remove();
                break;
            }
        };
    }
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
    const left = ((i - (i % 8)) / 8) * 150 + 450; // 左からの位置
    const top = (i % 8) * 50; // 上からの位置
    const width = 200;
    const height = 50;
    const textbox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, top, width, height);

    const textbox_content = textbox.getText();
    textbox_content.setText(message);

    const style = textbox_content.getTextStyle();
    style.setForegroundColor(getRandomRGB());
    style.setFontFamily("HiraginoSans-W3")
    style.setFontSize(20);
}

// ユーティリティ //////////////////////////////////////////////////////////////////////////////

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
    const hue = Math.floor(Math.random() * 360); // ランダムな色相
    const saturation = 100; // 最大彩度
    const brightness = 100; // 最大明度

    // RGBに変換
    const rgb = hsvToRgb(hue, saturation, brightness);

    // カラーコードを生成
    const rgb_code = `rgb(${rgb.r}, ${rgb.g}, ${rgb.b})`;
    const rgb_hex = rgbToHex(rgb);

    return rgb_hex;
}

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

function rgbToHex(rgb) {
    const r = rgb.r.toString(16).padStart(2, '0');
    const g = rgb.g.toString(16).padStart(2, '0');
    const b = rgb.b.toString(16).padStart(2, '0');

    return `#${r}${g}${b}`;
}
