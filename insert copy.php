<?php
//エラー表示
ini_set("display_errors", 1);

//1. POSTデータ取得
$tmp_name = $_FILES['excel_file']['tmp_name'];
$file_name = $_FILES['excel_file']['name'];

// ファイルの種類をチェック（セキュリティ対策）
$allowed_ext = array('xlsx', 'xls');
$ext = pathinfo($file_name, PATHINFO_EXTENSION);
if (!in_array($ext, $allowed_ext)) {
    // 許可されていないファイルタイプのとき
    echo '許可されていないファイルタイプです。';
    exit;
}

// ファイルを一時ディレクトリから目的の場所に移動（セキュリティ対策）
$upload_dir = 'uploads/'; // アップロード先のディレクトリ
$upload_file = $upload_dir.basename($file_name);
move_uploaded_file($tmp_name, $upload_file);

// ここからExcelファイルを読み込んでDBにINSERTする処理
include('./vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;

//2. DB接続します
try {
    //Password:MAMP='root',XAMPP=''
    $pdo = new PDO('mysql:dbname=avails_db;charset=utf8;host=localhost','root','');
} catch (PDOException $e) {
exit('DBError:'.$e->getMessage());
}

// Excelファイルを読み込む
$reader = new XlsxReader();
$reader->setReadDataOnly(true); // 読み込みモード
$spreadSheet = $reader->load($upload_file); // エクセルファイル読み込み
$highestRow = $spreadSheet->getActiveSheet()->getHighestDataRow();

$sheetData = $spreadSheet->getActiveSheet()->toArray(); // 全てのデータを配列で取得
$columnNames = $sheetData[1]; // 2行目のカラム名を取得

// DBのカラム名を配列に格納
$dbColumns = ['EntryType', 'WorkType', 'ALID', 'SeriesAltID', 'SeasonAltID', 'EpisodeAltID', 'SeriesContentID', 'SeasonContentID', 'EpisodeContentID', 'SeriesTitleInternalAlias', 'EpisodeTitleInternalAlias', 'EpisodeNumber', 'LicenseType', 'FormatProfile', 'LicenseRightsDescription', 'Start', 'End', 'SeasonNumber'];

// DBカラムに一致する列の値を配列に格納
$values = [];
foreach ($sheetData as $rowIndex => $row) {
    if ($rowIndex >= 3) { // 2行目以降の行を処理
        $rowValues = [];
        foreach ($dbColumns as $column) {
            $columnIndex = array_search($column, $columnNames);
            if ($columnIndex !== false) {
                $rowValues[] = $row[$columnIndex];
            }
        }
        $values[] = $rowValues;
    }
}

// $dbColumnsに基づいて$valuesの配列を並び替える
foreach ($values as &$row) {
    $row = array_combine($dbColumns, $row);
    $row = array_values($row); // 連想配列を数値インデックスの配列に戻す
}
// var_dump($values);

//３．データ登録SQL作成
$sql = 'INSERT INTO avails_an_table (FileName, UploadDate, ';
$sql .= implode(', ', $dbColumns);
$sql .= ') VALUES (?, ?, ';
$placeholders = array_fill(0, count($dbColumns), '?');
$sql .= implode(', ', $placeholders);
$sql .= ')';

// 4. データ登録処理
$stmt = $pdo->prepare($sql);
$stmt_update = $pdo->prepare("UPDATE avails_an_table SET FileName = ?, UploadDate = ?, EntryType = ?, WorkType = ?, ALID = ?, SeriesAltID = ?, SeasonAltID = ?, EpisodeAltID = ?, SeriesContentID = ?, SeasonContentID = ?, EpisodeContentID = ?, SeriesTitleInternalAlias = ?, EpisodeTitleInternalAlias = ?, EpisodeNumber = ?, LicenseType = ?, FormatProfile = ?, LicenseRightsDescription = ?, Start = ?, End = ?, SeasonNumber = ? WHERE ALID = ?");
$pdo->beginTransaction(); // トランザクション開始

$currentDate = date('Y-m-d H:i:s'); // 現在の日時をYYYY-MM-DD HH:II:SS形式で取得
foreach ($values as $row) {
    $alid = $row[2];

    // 重複チェック
    $check_sql = "SELECT COUNT(*) FROM avails_an_table WHERE ALID = ?";
    $check_stmt = $pdo->prepare($check_sql);
    $check_stmt->execute([$alid]);
    $count = $check_stmt->fetchColumn();

    if ($count > 0) {
        // 重複する場合: 更新処理
        $stmt_update->execute(array_merge([$file_name, $currentDate], $row, [$alid]));
    } else {
        // 重複しない場合: 挿入処理
        // var_dump(array_merge([$file_name, $currentDate], $row));
        $stmt->execute(array_merge([$file_name, $currentDate], $row));
    }
}

$pdo->commit(); // コミット

//４．データ登録処理後
if ($stmt->errorCode() !== '00000') {
    //SQL実行時にエラーがある場合（エラーオブジェクト取得して表示）
    $error = $stmt->errorInfo();
    exit("SQLError:".$error[2]);
    $error_update = $stmt_update->errorInfo();
    exit("SQLError_update:".$error_update[2]);
}else{

//５．index.phpへリダイレクト
    header("Location: index.php");
    exit();
}

// 不要なファイルを削除
unlink($upload_file);


//TODO: エクセルの日付フォーマットを文字列に変換して日付で登録
//TODO: MovieフォーマットのAvailsも取り込めるように = DBカラムの変更
//TODO: FullDeleteの概念これでいいんやっけ
//TODO: 空白行を無視
//TODO: 最後の行と列までを処理範囲とする
//FIXME: 最後の1行を読み込めない

?>