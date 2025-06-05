// 日本語カルチャー設定
GC.Spread.Common.CultureManager.culture("ja-jp");
//GC.Spread.Sheets.LicenseKey = "ここにSpreadJSのライセンスキーを設定します";

// SpreadJSの設定
document.addEventListener("DOMContentLoaded", () => {
    const previousButton = document.getElementById('previous');
    const nextButton = document.getElementById('next');
    const spread = new GC.Spread.Sheets.Workbook("ss");
    let reportSheet;
    //----------------------------------------------------------------
    // sjs形式のテンプレートシートを読み込んでレポートシートを実行します
    //----------------------------------------------------------------
    const res = fetch('reports/Fixed-Invoice.sjs').then((response) => response.blob())
        .then((myBlob) => {
            spread.open(myBlob, () => {
                console.log(`読み込み成功`);
                reportSheet = spread.getSheetTab(0);

                // レポートシートのオプション設定
                reportSheet.renderMode('PaginatedPreview');
                reportSheet.options.printAllPages = true;

                // レポートシートの印刷設定
                var printInfo = reportSheet.printInfo();
                printInfo.showBorder(false);
                printInfo.zoomFactor(1);
                reportSheet.printInfo(printInfo);
                initPage();
            }, (e) => {
                console.log(`***ERR*** エラーコード（${e.errorCode}） : ${e.errorMessage}`);
            });
        });


    //------------------------------------------
    // 前のページボタン押下時の処理
    //------------------------------------------    
    previousButton.onclick = function () {
        const page = reportSheet.currentPage();
        if (page != 0) {
            reportSheet.currentPage(page - 1);
            initPage()
        }
    }

    //------------------------------------------
    // 次のページボタン押下時の処理
    //------------------------------------------    
    nextButton.onclick = function () {
        const page = reportSheet.currentPage();
        if (page < reportSheet.getPagesCount() - 1) {
            reportSheet.currentPage(page + 1);
            initPage()
        }
    }

    //------------------------------------------
    // 現在のページと全ページ数の表示
    //------------------------------------------      
    function initPage() {
        document.getElementById('current').innerHTML = reportSheet.currentPage() + 1;
        document.getElementById('all').innerHTML = reportSheet.getPagesCount();
    }
});

