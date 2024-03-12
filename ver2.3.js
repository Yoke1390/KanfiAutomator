class Member {
    constructor(name) {
        this.name = name;
        this.note1 = name + "1枚目。このテキストは変更しないでください。";
        this.note2 = name + "2枚目。このテキストは変更しないでください。";
        this.nickname = name; // ニックネームの初期値を本名にする
        this.messages = [];
        this.slide1 = null;
        this.slide2 = null;
    }
}

function main() {
    // 感フィ君のスプレッドシートを取得。
    const mainSheet = connectMainSheet();
    // D12のセルにログを記録する関数を作成
    const log = createLogger(mainSheet.getRange("d12"));
    log("実行開始");

    // データの処理 ///////////////////////////////////////////////////
    log("アンケート結果の取得中...");
    const { memberNameSource, nickNamesSource, messageSource } = fetchKanfiData(mainSheet);
    log("メンバーの一覧を作成中...");
    const members = makeMembersMap(memberNameSource);
    log("読んでほしい名前を設定中...");
    setNickNames(nickNamesSource, members);
    log("メッセージを割り当て中...");
    setMessages(messageSource, members);

    // スライドの処理 /////////////////////////////////////////////////
    log("感フィを作成するスライドと接続中...");
    const kanfiSlide = connectKanfiSlide(mainSheet);
    log("既存のスライドを確認中...");
    assignExsistingSlidestoMembers(members, kanfiSlide);
    for (const member in members) {
        log(member.name + "の1枚目のスライドを処理中...");
        firstSlide(member, kanfiSlide);
        log(member.name + "の2枚目のスライドを処理中...");
        secondSlide(member, kanfiSlide);
    };
}

// データの処理 //////////////////////////////////////////////////////////////////////////////

function connectMainSheet() {
    // GoogleスプレッドシートのAPIを使ってアクティブなスプレッドシートを取得
    const spread = SpreadsheetApp.getActiveSpreadsheet();
    // スプレッドシート内の"メイン" という名前のシートを取得
    const mainSheet = spread.getSheetByName("メイン");
    return mainSheet;
}

function fetchKanfiData(mainSheet) {
    // アンケートデータが存在するスプレッドシートのURLを取得し、そのURLを使ってシートを取得
    // 感フィのスライドのURLを取得し、そのURLを使ってプレゼンテーションを開く
    const kanfiDataSheet = SpreadsheetApp.openByUrl(mainSheet.getRange("b3").getValue()).getActiveSheet();

    const memberNameSource = kanfiDataSheet.getRange(1, 4, 1, kanfiDataSheet.getLastColumn() - 3).getValues()[0];
    for (let i = 0; i < memberNameSource.length; i++) {
        memberNameSource[i] = deleteSpace(memberNameSource[i]);
    }

    const nickNamesSource = kanfiDataSheet.getRange(2, 2, kanfiDataSheet.getLastRow() - 1, 2).getValues();

    const messageSource = kanfiDataSheet.getRange(2, 4, kanfiDataSheet.getLastRow() - 1, kanfiDataSheet.getLastColumn() - 3).getValues();

    return { memberNameSource, nickNamesSource, messageSource };
}

function makeMembersMap(memberNameSource) {
    const members = {};
    for (const name of memberNameSource) {
        members[name] = new Member(name);
    };
    return members;
}

function setNickNames(nickNamesSource, members) {
    for (const row of nickNamesSource) {
        const name = deleteSpace(row[0]); // 回答に記載されていた本名から空白を削除
        members[name].nickname = row[1];
    };
    // 正しいニックネームが設定されているか確認
    // for (const name in members) {
    //     console.log(members[name].nickname); // テスト用なので、コンソールに出力すればOK
    // };
}

function setMessages(messageSource, members) {
    let i = 0;
    for (const name in members) {
        members[name].messages = messageSource.map(function (row) {
            return deleteSpace(row[i]);
        }).filter(function (message) {
            return message !== "";
        });
        i++;
    }
    // 正しいメッセージが設定されているか確認するためのコード。main関数の中に記述する。
    // for (const name in members) {
    //     console.log(members[name].messages); // テスト用なので、コンソールに出力すればOK
    // };
}

// スライドの処理 //////////////////////////////////////////////////////////////////////////////

function connectKanfiSlide(mainSheet) {
    const kanfiSlide = SlidesApp.openByUrl(mainSheet.getRange("d3").getValue());
    return kanfiSlide;
}

function assignExsistingSlidestoMembers(members, kanfiSlide) {

    // 各メンバーのスライドの識別子を格納する集合を作成
    setOfNotes = new Set();
    for (const name in members) {
        setOfNotes.add(members[name].note1);
        setOfNotes.add(members[name].note2);
    }

    const exsistingSlides = kanfiSlide.getslides();
    const slidesToRemove = []; //  todo：削除リストが必要かどうかテスト
    for (const slide of exsistingSlides) {
        const note = deleteSpace(slide.getnotespage().getspeakernotesshape().gettext().asstring()); // スライドのスピーカーノートを取得
        if (memberNotes.includes(note)) {
            // スピーカーノートが識別子リストに含まれている場合は、スライドを保存
            //  TODO：メンバーに既存のスライドを割り当てるアルゴリズムを変更。スピーカーノートのテキストを解析して、誰の何枚目のスライドかを判断する。
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

function firstSlide(member, kanfiSlide) {
    // TODO: 1枚目のスライドを作成する関数を作成
    if (member.slide1 !== null) {
        checkNicknameOnSlide(member.name, member.nickname, member.slide1);
        return;
    }
}

function secondSlide(member, kanfiSlide) { }

function checkNicknameOnSlide(name, nickname, slide) {
    // TODO: スライドにニックネームが正しく表示されているか確認する関数を作成
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
    const rgbCode = `rgb(${rgb.r}, ${rgb.g}, ${rgb.b})`;

    return rgbCode;
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
