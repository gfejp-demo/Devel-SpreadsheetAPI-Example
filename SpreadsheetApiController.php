<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Libs\GoogleDrive;

class DevelSpreadsheetController extends Controller
{

    protected $csvSpreadsheetId;

    /**
    * スプレッドシートの作成
    * $client: Googleクライアント
    */
    public function develOpenSheet($client) {

        try {
            // スプレッドシートサービス オブジェクトを生成
            $service = new \Google_Service_Sheets($client);

            // スプレッドシートを作成
            $postBody = new \Google_Service_Sheets_Spreadsheet([
                'properties' => [
                    'title' => 'New Spreadsheet 1' // スプレッドシート名
                ]
            ]);

            $optParams = array();
            $spreadsheet = $service->spreadsheets->create($postBody, $optParams);
            $spreadsheetId = $spreadsheet->spreadsheetId;

            // 新しいシートを追加
            $body = new \Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => [
                    'addSheet' => [
                        'properties' => [
                            'title' => 'New Sheet 1' // シート名
                        ]
                    ]
                ]
            ]);

            $response = $service->spreadsheets->batchUpdate($spreadsheetId, $body);

        } catch(\Exception $e) {
            $msg = $e -> getMessage();
            return $msg;
        }

    }


    /**
    * スプレッドシートへのデータの書き出し
    * $client: Googleクライアント
    */
    public function develExportCsv($client) {

        // 出力するデータ
        $values = [
            ["id","class","email","name","math","science","english"],
            ["1","1A","john@example.com","John J. Coons","90","88","96"],
            ["2","1A","crystal@example.com","Crystal C. Burnett","32","44","89"],
            ["3","1B","anthony@example.com","Anthony T. Dudley","72","68","24"],
            ["4","1B","francisca@example.com","Francisca H. Rapp","89","94","92"]
        ];
        $range = 'CSV 1!A1:G5';

        try {
            // スプレッドシートサービス オブジェクトを生成
            $service = new \Google_Service_Sheets($client);

            // スプレッドシートを作成
            $postBody = new \Google_Service_Sheets_Spreadsheet([
                'properties' => [
                    'title' => 'CSV import 1' // スプレッドシート名
                ]
            ]);

            $optParams = array();
            $spreadsheet = $service->spreadsheets->create($postBody, $optParams);
            $spreadsheetId = $spreadsheet->spreadsheetId;
            $this->csvSpreadsheetId = $spreadsheetId;

            // 新しいシートを追加
            $body = new \Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => [
                    'addSheet' => [
                        'properties' => [
                            'title' => 'CSV 1'
                        ]
                    ]
                ]
            ]);
            $response = $service->spreadsheets->batchUpdate($spreadsheetId, $body);

            // 出力するデータを準備
            $body = new \Google_Service_Sheets_ValueRange([
                'values' => $values
            ]);
            $body->setValues($values);

            $params = ['valueInputOption' => 'USER_ENTERED'];

            // データを出力
            $result = $service->spreadsheets_values->update(
                $spreadsheetId, // スプレッドシートID
                $range, // 出力先のデータレンジ
                $body, // 出力データ
                $params // 出力オプション
            );

        } catch(\Exception $e) {
            $msg = $e -> getMessage();
            return $msg;
        }

    }


    /**
    * スプレッドシートの書式設定
    * $client: Googleクライアント
    */
    public function develFormatSheet($client) {

        try {
            // ドライブサービス オブジェクトを生成
            $driveClient = new \Google_Service_Drive($client);

            // 名前を指定してドライブ内のスプレッドシートを検索
            $result = $driveClient->files->listFiles([
                "q" => "name='CSV import 1'"
            ]);
            $file = $result->getFiles()[0];
            $spreadsheet_id = $file->getId();

            // スプレッドシートサービス オブジェクトを生成
            $spreadsheet_service = new \Google_Service_Sheets($client);

            // シートIDを取得
            $sheet_id;
            $response = $spreadsheet_service->spreadsheets->get($spreadsheet_id);
            $sheets = $response->getSheets();
            foreach ($sheets as $sheet) {
                $properties = $sheet->getProperties();
                $sheet_title = $properties->getTitle();

                // 名前の一致するシートを検索
                if ($sheet_title == "CSV 1") {
                    $sheet_id = $properties->getSheetId();
                    break;
                }
            }

            // 書式を設定する範囲と書式を準備
            $request_data = [
                'repeatCell' => [
                    'fields' => 'userEnteredFormat(backgroundColor)',
                    'range' => [
                        'sheetId' => $sheet_id,
                        'startRowIndex' => 0, // 行の開始位置
                        'endRowIndex' => 5, // 行の終了位置
                        'startColumnIndex' => 0, // 列の開始位置
                        'endColumnIndex' => 7, // 列の終了位置
                    ],
                    'cell' => [
                        'userEnteredFormat' => [
                            'backgroundColor' => [  // RGB値でセルの背景色を指定
                                'red' => 234/255,
                                'green' => 143/255,
                                'blue' => 143/255
                            ]
                        ],
                    ],
                ],
            ];
            $requests = [new \Google_Service_Sheets_Request($request_data)];

            // 書式を設定
            $batchUpdateRequest = new \Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => $requests
            ]);
            $response = $spreadsheet_service->spreadsheets->batchUpdate($spreadsheet_id, $batchUpdateRequest);

        } catch(\Exception $e) {
            $msg = $e -> getMessage();
            return $msg;
        }

    }


    /**
    * ピボットテーブルの生成
    * $client: Googleクライアント
    */
    public function develGeneratePivot($client) {
        try {
            // ドライブサービス オブジェクトを生成
            $driveClient = new \Google_Service_Drive($client);

            // 名前を指定してドライブ内のスプレッドシートを検索
            $result = $driveClient->files->listFiles([
                "q" => "name='CSV import 1'"
            ]);
            $file = $result->getFiles()[0];
            $spreadsheet_id = $file->getId();

            // スプレッドシートサービス オブジェクトを生成
            $spreadsheet_service = new \Google_Service_Sheets($client);

            // シートIDを取得
            $sheet_id;
            $response = $spreadsheet_service->spreadsheets->get($spreadsheet_id);
            $sheets = $response->getSheets();
            foreach ($sheets as $sheet) {
                $properties = $sheet->getProperties();
                $sheet_title = $properties->getTitle();

                // 名前の一致するシートを検索
                if ($sheet_title == "CSV 1") {
                    $sheet_id = $properties->getSheetId();
                    break;
                }
            }

            // ピボットテーブルの設定を準備
            $pt_requests = [
                'updateCells' => [
                    'rows' => [
                        'values' => [
                            [
                                'pivotTable' => [
                                    'source' => [
                                        'sheetId' => $sheet_id,
                                        'startRowIndex' => 0, // 行の開始位置
                                        'startColumnIndex' => 0, // 列の開始位置
                                        'endRowIndex' => 5, // 行の終了位置
                                        'endColumnIndex' => 7 // 列の終了位置
                                    ],
                                    'rows' => [
                                        [
                                            'sourceColumnOffset' => 1,
                                            'sortOrder' => 'ASCENDING',
                                            'showTotals' => true,
                                        ],
                                        [
                                            'sourceColumnOffset' => 2,
                                            'sortOrder' => 'ASCENDING',
                                            'showTotals' => true,
                                        ]
                                    ],
                                    'values' => [
                                        [
                                            'summarizeFunction' => 'AVERAGE', // 平均値を出力
                                            'sourceColumnOffset' => 4
                                        ],
                                        [
                                            'summarizeFunction' => 'AVERAGE', // 平均値を出力
                                            'sourceColumnOffset' => 5
                                        ],
                                        [
                                            'summarizeFunction' => 'AVERAGE', // 平均値を出力
                                            'sourceColumnOffset' => 6
                                        ]
                                    ],
                                ]
                            ]
                        ]
                    ],
                    'start' => [
                        'sheetId' => $sheet_id,
                        'rowIndex' => 7,
                        'columnIndex' => 2
                    ],
                    'fields' => 'pivotTable'
                ]
            ];

            // ピボットテーブルを生成
            $batchUpdateRequest = new \Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => $pt_requests
            ]);
            $response = $spreadsheet_service->spreadsheets->batchUpdate($spreadsheet_id, $batchUpdateRequest);

        } catch(\Exception $e) {
            $msg = $e->getMessage();
            return $msg;
        }

    }



    public function devel_spreadsheet_opensheet($client) {
        $msg = $this->develSpreadsheetOpensheet($client);
        return view('ok')->with('msg', $msg);
    }

    public function devel_spreadsheet_exportcsv($client) {
        $msg =$this->develSpreadsheetExportcsv($client);
        return view('ok')->with('msg', $msg);
    }

    public function devel_spreadsheet_formatsheet($client) {
        $msg =$this->develSpreadsheetFormatsheet($client);
        return view('ok')->with('msg', $msg);
    }

    public function devel_spreadsheet_generate($client) {
        $msg =$this->develSpreadsheetGenerate($client);
        return view('ok')->with('msg', $msg);
    }

}
