function create() {
    const spread = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = spread.getSheetByName("メイン");
    const logdata = mainSheet.getRange("d12");
    logdata.setValue("実行開始");

    const sheet = SpreadsheetApp.openByUrl(mainSheet.getRange("b3").getValue()).getActiveSheet();
    logdata.setValue("アンケート結果の取得中...");

    var presentation = SlidesApp.openByUrl(mainSheet.getRange("d3").getValue());
    logdata.setValue("スライドの取得中...");
    Logger.log(presentation.getName());

    while (presentation.getSlides().length != 0) {
        presentation.getSlides()[0].remove();
        logdata.setValue("スライドの初期化中...");
    }


    var n = 0;
    while (sheet.getRange(1, n + 5).getValue().replace("\n", "") != "") {
        logdata.setValue("アンケート結果の処理/出力中..." + (n + 1) + "人目");

        presentation.appendSlide();
        var slide = presentation.getSlides()[2 * n];

        var personName = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 155, 150, 400, 100);
        var textPN = personName.getText();
        textPN.setText(sheet.getRange(1, n + 5).getValue().replace("\n", ""));
        textPN.getTextStyle().setForegroundColor("#ffffff").setFontSize(60);
        textPN.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

        presentation.appendSlide();
        var slide = presentation.getSlides()[2 * n + 1];

        var personName = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 155, 150, 400, 100);
        var textPN = personName.getText();
        textPN.setText(sheet.getRange(1, n + 5).getValue().replace("\n", ""));
        textPN.getTextStyle().setForegroundColor("#ffffff").setFontSize(60);
        textPN.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

        var count = 1;
        while (sheet.getRange(count, 1).getValue() != "") {
            count = count + 1;
        }

        var sgn = 0;
        for (var i = 0; i < count; i++) {
            var data = sheet.getRange(i + 2, n + 5).getValue();
            if (data.trim() == "") continue;
            var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, ((sgn - (sgn % 8)) / 8) * 150 + 450, (sgn % 8) * 50, 200, 50);
            var textRNG = shape.getText();
            textRNG.setText(data);
            textRNG.getTextStyle().setForegroundColor(225 - Math.ceil(Math.random() * 4) * 30, 225 - Math.ceil(Math.random() * 4) * 30, 225 - Math.ceil(Math.random() * 4) * 30).setFontSize(20);
            sgn++;
        }
        n++;
    }

    logdata.setValue("処理完了/次のデータ受付中");
}

function reverse() {
    const spread = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = spread.getSheetByName("メイン");
    const logdata = mainSheet.getRange("d12");
    logdata.setValue("色反転実行開始");
    logdata.setValue("スライドの取得中...");
    var presentation = SlidesApp.openByUrl(mainSheet.getRange("d17").getValue());
    Logger.log(presentation.getName());
    var slides = presentation.getSlides()
    for (var i = 0; i < slides.length; i++) {
        logdata.setValue("色反転中..." + (i + 1) + "枚目");
        var slide = slides[i];
        var shapes = slide.getShapes();
        for (var j = 0; j < shapes.length; j++) {
            var shape = shapes[j];
            if (shape.getText().asString().trim() == "") continue;
            var style = shape.getText().getTextStyle();
            if (style.getForegroundColor() == null) continue;
            var R = style.getForegroundColor().asRgbColor().getRed();
            var G = style.getForegroundColor().asRgbColor().getGreen();
            var B = style.getForegroundColor().asRgbColor().getBlue();
            var maxRGB = Math.max(R, G, B) + Math.min(R, G, B)

            style.setForegroundColor(maxRGB - R, maxRGB - G, maxRGB - B)

        }

    }

    logdata.setValue("色反転完了");
}
