// =====================================================================
// summary_dashboard.gs
// 売上データを月次集計し、サマリーシートへ書き込み＋棒グラフを生成する
// =====================================================================

var SALES_SHEET_NAME   = '売上データ';
var SUMMARY_SHEET_NAME = '月次サマリー';
var CHART_TITLE        = '月次売上推移';

// ---------------------------------------------------------------------
// メイン処理：集計 → サマリー書き込み → グラフ更新
// 毎朝9時のトリガーから自動呼び出しされる
// ---------------------------------------------------------------------
function updateSummaryDashboard() {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var salesSheet   = ss.getSheetByName(SALES_SHEET_NAME);
  var summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);

  // 「月次サマリー」シートが存在しない場合は新規作成
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SUMMARY_SHEET_NAME);
  }

  // 売上データを集計する
  var monthlyData = aggregateSalesByMonth(salesSheet);

  // サマリーシートを更新する
  writeSummary(summarySheet, monthlyData);

  // 棒グラフを更新する
  updateBarChart(ss, summarySheet);

  Logger.log('サマリーダッシュボードの更新が完了しました。');
}

// ---------------------------------------------------------------------
// 「売上データ」シートを読み込み、月ごとに合計売上と件数を集計する
// 戻り値: [{ month, label, total, count }, ...] ※月の昇順でソート済み
// ---------------------------------------------------------------------
function aggregateSalesByMonth(salesSheet) {
  var lastRow = salesSheet.getLastRow();

  // ヘッダー行を除いたデータが存在しない場合は空配列を返す
  if (lastRow < 2) {
    return [];
  }

  // A列（日付）・D列（金額）のみ取得してメモリを節約
  var dateValues   = salesSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var amountValues = salesSheet.getRange(2, 4, lastRow - 1, 1).getValues();

  var map = {}; // キー: 'YYYY-MM'、値: { label, total, count }

  for (var i = 0; i < dateValues.length; i++) {
    var rawDate = dateValues[i][0];
    var amount  = amountValues[i][0];

    // 日付が空またはゼロの行はスキップ
    if (!rawDate || rawDate === '') continue;

    var date = new Date(rawDate);
    var year  = date.getFullYear();
    var month = date.getMonth() + 1; // 0始まりを1始まりに変換

    var key   = year + '-' + ('0' + month).slice(-2);
    var label = year + '年' + month + '月';

    if (!map[key]) {
      map[key] = { label: label, total: 0, count: 0 };
    }

    map[key].total += Number(amount) || 0;
    map[key].count += 1;
  }

  // キー（YYYY-MM）の昇順でソートして配列に変換
  var keys = Object.keys(map).sort();
  return keys.map(function(key) {
    return {
      month: key,
      label: map[key].label,
      total: map[key].total,
      count: map[key].count
    };
  });
}

// ---------------------------------------------------------------------
// 「月次サマリー」シートをクリアしてヘッダー＋集計データを書き込む
// ---------------------------------------------------------------------
function writeSummary(summarySheet, monthlyData) {
  // シートの内容を全消去（書式も含めてリセット）
  summarySheet.clearContents();

  // ヘッダー行を設定
  var header = [['月', '合計売上', '件数']];
  summarySheet.getRange(1, 1, 1, 3).setValues(header);
  summarySheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  if (monthlyData.length === 0) {
    Logger.log('集計対象のデータがありません。');
    return;
  }

  // データ行を一括書き込み
  var rows = monthlyData.map(function(d) {
    return [d.label, d.total, d.count];
  });
  summarySheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // B列（合計売上）を通貨形式にフォーマット
  summarySheet.getRange(2, 2, rows.length, 1).setNumberFormat('¥#,##0');
}

// ---------------------------------------------------------------------
// 「月次サマリー」シートのデータをもとに棒グラフを作成・更新する
// 既存のグラフが同名タイトルで存在する場合は削除して再作成する
// ---------------------------------------------------------------------
function updateBarChart(ss, summarySheet) {
  var lastRow = summarySheet.getLastRow();

  // ヘッダー行のみ（データなし）の場合はグラフを作らない
  if (lastRow < 2) {
    Logger.log('グラフ作成に必要なデータがありません。');
    return;
  }

  // 既存の同タイトルグラフを削除
  var existingCharts = summarySheet.getCharts();
  existingCharts.forEach(function(chart) {
    if (chart.getOptions().get('title') === CHART_TITLE) {
      summarySheet.removeChart(chart);
    }
  });

  // グラフ用データ範囲（月ラベル＋合計売上）
  var dataRange = summarySheet.getRange(1, 1, lastRow, 2);

  var chart = summarySheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setOption('title', CHART_TITLE)
    .setOption('hAxis.title', '月')
    .setOption('vAxis.title', '売上金額（円）')
    .setOption('legend.position', 'none')
    .setOption('width', 700)
    .setOption('height', 420)
    .setPosition(2, 5, 0, 0) // 2行目・5列目の位置にグラフを配置
    .build();

  summarySheet.insertChart(chart);
  Logger.log('棒グラフを更新しました。');
}

// ---------------------------------------------------------------------
// 毎朝9時の時間駆動トリガーを登録する
// ※ GASエディタでこの関数を一度だけ手動実行してください
//    重複登録を防ぐため、同名の既存トリガーは削除してから再登録します
// ---------------------------------------------------------------------
function setupDailyTrigger() {
  var targetFunctionName = 'updateSummaryDashboard';

  // 同名の既存トリガーをすべて削除して重複を防ぐ
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === targetFunctionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎朝9時〜10時の間に実行されるトリガーを登録
  ScriptApp.newTrigger(targetFunctionName)
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log('毎朝9時のトリガーを登録しました。');
}
