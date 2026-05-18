# 変換エンジン比較検証

Markdown → docx 変換における 3 方式の比較検証記録（v0.5.0 時点）。
v0.3.0 / v0.4.0 で得た知見をもとに、現行方式の継続か .NET ライブラリ内包方式への
移行かを判断する材料を提供する。本ドキュメントは **比較・推奨** までであり、
v0.5.0 時点では正式採用は行わない。

---

## 1. 背景

v0.3.0 で発生した本文消失問題は、Word COM の段落再計算・リスト書式伝播に起因していた。
v0.4.0 で docx XML 検査による回帰テスト基盤を整備した結果、現行方式の安定性は
検証可能になったが、以下の課題が残る。

- 統合テストに Microsoft Word が必要 → CI で本体実行できない
- Word の自動補正・伝播挙動への依存が残る（リスト処理など）
- 起動オーバーヘッドが大きい（Word プロセス起動 + COM Marshal）
- Word のライセンス前提が変換要件に含まれる

これらは「Word を使わない .NET ライブラリ完結型」へ移行することで根治可能。
ただし移行コスト・出力品質の差異は未検証のため、本ドキュメントで整理する。

---

## 2. 比較対象

外部 EXE の同梱は方針外のため、以下 3 方式を対象とする。

| 方式 | 構成 |
|------|------|
| A. 現行 Word COM 方式 | `Microsoft.Office.Interop.Word`（実機 Word 必須） |
| B. Markdig + Open XML SDK | `Markdig`（パース）+ `DocumentFormat.OpenXml`（書き出し） |
| C. Markdig + HTML + HtmlToOpenXml | `Markdig`（HTML 化）+ `HtmlToOpenXml`（docx 化） |

### 方式 A: 現行 Word COM 方式

- 実装位置: `src/md2doc/WordInteropConverter.cs`
- 4 フェーズ方式（ParseBlocks → WriteAllText → ApplyStyles → InsertTables）で
  v0.3.0 の本文消失問題を回避済み
- 実行時に Microsoft Word を起動して COM 経由で docx を生成

### 方式 B: Markdig + Open XML SDK

- `Markdig`（[GitHub](https://github.com/xoofx/markdig), BSD-2-Clause）で
  Markdown を AST に変換
- `DocumentFormat.OpenXml`（[GitHub](https://github.com/dotnet/Open-XML-SDK), MIT）で
  OOXML（docx 内部の XML）を直接構築
- AST を走査し、各ノードに対応する OOXML 要素（`Paragraph` / `Run` / `Table` 等）を
  手動でマッピングする

### 方式 C: Markdig + HTML + HtmlToOpenXml

- `Markdig` で Markdown を HTML に変換（標準機能）
- `HtmlToOpenXml`（[GitHub](https://github.com/onizet/html2openxml), MIT）で
  HTML フラグメントを OOXML に変換し docx に挿入
- Markdown → HTML → docx の 2 段変換パイプライン

---

## 3. 比較表

### 機能・品質

| 観点 | A. Word COM | B. Open XML SDK | C. HtmlToOpenXml |
|------|------------|----------------|-----------------|
| 見出し（H1〜H3） | ◎ Word 標準スタイル | ◯ `ParagraphStyleId="Heading1"` で対応 | ◯ `<h1>` → Heading スタイル自動マップ |
| 通常段落 | ◎ | ◎ | ◎ |
| 箇条書き（`-`） | △ 自動伝播あり、2 パス補正必要 | ◯ Numbering 定義を明示構築 | ◯ `<ul>` → 箇条書きスタイル |
| 順序付きリスト | × 未対応（v0.3.0 時点で通常段落扱い） | ◯ Numbering 定義で対応可 | ◯ `<ol>` → 番号リスト |
| 表 | ◎ Word ネイティブ | ◯ `Table`/`TableRow`/`TableCell` 構築 | ◯ `<table>` を変換 |
| 改ページ | ◎ `\f` 一文字で挿入 | ◯ `<w:br w:type="page"/>` 要素 | △ CSS `page-break-before` 経由（要検証） |
| 日本語本文 | ◎ | ◎ UTF-8 XML | ◎ |
| インライン書式（太字・斜体・コード） | △ マーカー除去のみ（装飾なし） | ◯ `RunProperties` で完全対応可 | ◯ HTML `<strong>`/`<em>`/`<code>` をマップ |
| Word での再編集 | ◎ | ◎ | ◎ |

### 開発・運用

| 観点 | A. Word COM | B. Open XML SDK | C. HtmlToOpenXml |
|------|------------|----------------|-----------------|
| 必要ランタイム（実行時） | Microsoft Word 必須 | .NET 8 Runtime のみ | .NET 8 Runtime のみ |
| 必要ライブラリ | Office Interop（PIA / NetOffice 等） | DocumentFormat.OpenXml | HtmlToOpenXml + Markdig + DocumentFormat.OpenXml |
| 配布物サイズ増加 | なし（Word 前提） | +約 5〜6 MB | +約 7〜8 MB |
| 変換速度 | 遅い（Word 起動 3〜5 秒 + COM オーバーヘッド） | 高速（XML 直書き、〜数百ms） | 高速（〜1 秒程度） |
| CI 統合テスト | △ Word 必須、現状スキップ表示 | ◎ Linux/Mac/Windows どこでも実行可 | ◎ 同左 |
| 既存 `DocxInspector` 再利用 | ◎（v0.4.0 時点で利用中） | ◎ そのまま流用可 | ◎ そのまま流用可 |
| 実装複雑度 | 中（4 フェーズ + COM 安定化） | 高（OOXML を手書き構築） | 低（HTML 経由の薄いパイプライン） |
| デバッグ容易性 | △ COM 例外・遅延再計算でトラブル多い | ◯ XML を直接検査可能 | ◯ 中間 HTML を確認可能 |

### ライセンス・配布

| 観点 | A. Word COM | B. Open XML SDK | C. HtmlToOpenXml |
|------|------------|----------------|-----------------|
| ライブラリライセンス | Office Interop: 再配布可 | MIT | MIT（Markdig: BSD-2-Clause、OpenXml: MIT） |
| 利用者側ライセンス要件 | Microsoft Word が必要 | 不要 | 不要 |
| 庁内・閉域環境配布 | × Word 必須が制約 | ◎ DLL 同梱で完結 | ◎ DLL 同梱で完結 |
| ライセンス審査の手間（庁内配布想定） | Word 利用前提を明記 | MIT のみ、低リスク | MIT + BSD-2-Clause、低リスク |

> ※ 各ライブラリの最新ライセンスは採用前に NuGet パッケージのメタデータで再確認すること。

---

## 4. 各方式の詳細評価

### 方式 A: 現行 Word COM の評価

**強み**
- 出力が本物の Word ファイル（テンプレートとの完全な整合）
- 見出し・表・改ページが Word ネイティブで自然
- 既に v0.3.0 / v0.4.0 で動作実績あり

**弱み**
- Word ライセンスが実行要件 → 庁内配布で「Word 必須」が制約となる
- CI で統合テストを動かせない（v0.4.0 でも passed 偽装の懸念）
- Word の自動補正・伝播に依存し、回帰再発リスクが残る
- 起動が遅く、大量変換のユースケースで不利
- 順序付きリスト未対応

### 方式 B: Markdig + Open XML SDK の評価

**強み**
- Word 不要 → 配布が docx 一枚で完結、庁内 PC で Office なしでも変換可能
- MIT ライセンスのみ → 庁内ライセンス審査が通りやすい
- OOXML を直接構築するため挙動が確定的（自動補正の罠なし）
- CI で完全自動テスト可能（`DocxInspector` がそのまま使える）
- v0.3.0 で発生した「Word の自動補正による消失・伝播」は構造的に起こらない

**弱み**
- OOXML の構築コードが冗長（見出し・段落・リスト・表・改ページごとに手書き）
- Numbering 定義（リスト書式）は別パートとして定義する必要があり、初期実装に手間
- 開発コストはほぼゼロからの再実装

**実装規模の目安**
- 既存 `MarkdownParser` を流用すれば、レンダラ部分は新規 300〜500 行程度
- Markdig 採用なら Markdown 解析は不要になり、内部モデルを Markdig AST に差し替える設計も可能

### 方式 C: Markdig + HtmlToOpenXml の評価

**強み**
- 実装が最も薄い（Markdown → HTML → docx の 2 段）
- Word 不要、MIT ライセンス
- HTML の意味論（`<h1>` / `<ul>` / `<table>`）と docx スタイルのマッピングが
  ライブラリ側で実装済み

**弱み**
- 出力品質が HtmlToOpenXml の HTML→OOXML マッピング実装に依存
- 改ページや細かい書式の指定方法が間接的（CSS 経由）になり、検証コストが高い
- 中間 HTML のデバッグが必要になる場面が出る
- メジャー Microsoft 製ではないため、長期メンテナンス性は OpenXml SDK より弱い

---

## 5. v0.5.0 時点での推奨

### 推奨: 方式 B（Markdig + Open XML SDK）を本命候補とする

理由:

1. **配布要件の改善**: Word 必須が外れ、庁内・閉域環境で Office なし端末でも変換可能
2. **テスト容易性**: CI でフル統合テストが回せ、回帰検出が確実
3. **挙動の確定性**: Word の自動補正・伝播による偶発バグが構造的に消える
4. **長期メンテナンス性**: Microsoft 公式 SDK で MIT、依存リスクが低い
5. **ライセンス審査**: MIT 単独で完結し、庁内利用判断が単純

### 補完案: 方式 C は POC 段階で並行評価

方式 B の実装コストが想定より高い場合、方式 C を「軽量代替」として並行評価する。
ただし出力品質の確認が必須。

### 方式 A は当面温存

v0.5.0 時点で削除はしない。以下の理由で互換モードとして残す。

- 既存ユーザーが現行出力に依存している可能性
- Word の標準テンプレートと完全に整合する利点
- 切り替え期間中の比較基準として有用

---

## 6. 採用判断のための次ステップ（v0.5.1 以降想定）

本ドキュメントを採用判断につなげるための提案：

1. **方式 B の POC 実装**
   - 別ブランチで Markdig + Open XML SDK 版のレンダラを試作
   - 既存 `MarkdownParser` か Markdig AST どちらをモデルにするか決定
   - 必須 Markdown 要素（見出し・段落・箇条書き・順序付き・表・改ページ）を対応

2. **既存テスト基盤での比較検証**
   - `DocxInspector` + `DocxConversionTests` を方式 B の出力にも適用
   - 同じ入力 Markdown で方式 A / B の出力 docx XML を比較

3. **実機 Word での目視確認**
   - 方式 B 出力を実機 Word で開き、編集容易性・スタイル整合を確認
   - スクリーンショット比較を `docs/` に記録

4. **エンジン切替の UI/設定対応の検討**
   - 切替を「設定で選択」「ビルド時固定」「段階的切替」のどれにするか方針決定

5. **v0.6.0 以降で方式 A 廃止判断**
   - 方式 B が品質・実用性で問題ないと判断できた時点で方式 A を deprecate

---

## 7. 対象外（v0.5.0 で扱わないこと）

- 現行エンジン（方式 A）の削除
- UI へのエンジン切替機能追加
- Pandoc 等の外部 EXE 同梱方式の検証
- Aspose / Spire 等の有償ライブラリ前提の実装
- 方式 B / C の本実装（v0.5.0 は比較記録まで）

---

## 8. v0.5.1〜v0.5.6 実装状況（方式 B）

### v0.5.1 POC 結果

方式 B（Markdig + Open XML SDK）の POC 実装を行い、基本要素の変換が実用可能であることを確認した。

### v0.5.2 本実装候補化

インライン書式（太字・斜体・コード）・ヘッダー・フッター・水平線・ソフトリターンを追加実装。
26 テスト（全 Word 不要）で品質を検証済み。詳細は `docs/poc-openxml.md` を参照。

### v0.5.3 変換方式選択 UI の追加

画面上で Open XML 方式 / Word COM 方式を切り替えられるようにした。
**Open XML 方式を初期値** として設定し、v1.0.0 に向けた標準候補として実機比較検証を開始。

### v0.5.4 実機確認結果（v0.5.3 実施分）

| 確認項目 | 方式 A（Word COM） | 方式 B（Open XML） |
|---------|------------------|------------------|
| 本文消失（v0.3.0 既知問題） | — | 再現しなかった ✅ |
| 段落番号付与不備 | — | 再現しなかった ✅ |
| 変換速度 | 約 10 秒 | 1 秒未満（体感 0.1 秒） ✅ |
| 移行可否 | — | 移行可能と判断できる見込み ✅ |

### v0.5.5 コード品質整理

ソースコードベースのレビューで検出したバグ修正・リファクタリングを実施。  
外部仕様に変更なし。

### v0.5.6 方式 B 標準採用の決定

v0.5.4 の実機確認・v0.5.5 の品質整備を経て、**方式 B（Open XML 方式）を標準方式として採用する** 判断を確定した。

- 本文消失・段落番号付与の不備：Open XML 方式では再現しなかった
- 変換速度：Word COM 方式 約 10 秒 → Open XML 方式 1 秒未満（体感 0.1 秒）
- テスト品質：Open XML 方式の 26 テストが CI で完全実行（Word 不要）
- ライセンス：MIT のみで庁内ライセンス審査が単純

**方式 A（Word COM 方式）は削除せず、互換確認用・退避用として引き続き維持する。**  
v0.6.0 以降は Open XML 方式を前提とした配布前検証フェーズとして進める。

---

## 9. 参考リンク

- Markdig: https://github.com/xoofx/markdig
- Open XML SDK: https://github.com/dotnet/Open-XML-SDK
- HtmlToOpenXml: https://github.com/onizet/html2openxml
- 関連ドキュメント:
  - `docs/dev-notes-word-interop.md` — 方式 A の COM 実装メモ
  - `docs/proposal-next-version.md` — 後継版設計提案
  - `docs/poc-openxml.md` — 方式 B の実装記録（v0.5.1〜v0.5.3）
  - `docs/test_scenarios.md` — 既存テスト基盤（方式 B / C でも流用可）
