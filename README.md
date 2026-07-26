# haribow-lighting-editor

HARIBOW演目の照明キューを、波形を見ながらドラッグで微調整できるWebエディタ。

## 背景

Roar復活の照明キューシート作業で、「口頭で秒数指定→Claude Codeがpythonのキュー修正→previzレンダー待ち→確認」という往復が非効率だったため、誰でも使える微調整エディタとして独立させたプロジェクト。

## ワークフロー

1. **インプット収集**: 音源ファイル＋既に決まっている部分（何秒から何の照明がいいか）
2. **AI一気出し**: Claude Codeが音楽解析込みで完全同期の照明案を作り、キューデータストア（Sheet）に登録
3. **エディタで微調整**: 担当者がこの画面で波形を見ながらキュー点をドラッグ調整
4. **反映**: 編集結果を正典（previzスクリプト・キューシートMarkdown）に反映

## 構成

- `index.html` — エディタ本体（静的、GitHub Pagesで配信）
- `gas/` — バックエンド（GAS、キューデータの読み書き先はGoogle Sheets）

## データ

キューデータの正本はGoogle Sheets（GASのScript Propertiesに`CUE_SS_ID`で指定）。
