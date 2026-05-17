# Word COM Interop 開発メモ

md2doc における Word COM Interop 実装上の知見・既知の制約・設計判断をまとめる。
v0.3.0 実装時の調査結果をもとに記述（2026-05-17）。

---

## 背景: v0.3.0 で発生した段落消失問題

### 症状

Markdown の見出し直後に来る段落・箇条書きの 1 行目が Word 出力に現れない場合があった。

```markdown
## 起動方法
- アプリを起動する          ← 出力されない
- Markdownファイルを選択する
```

さらに修正試行を重ねるうちに、箇条書きの偶数行が消える / 奇数行のみ残る という
交互消失パターンも観測された。

### 根本原因

旧実装は「段落を 1 行追加 → スタイル適用 → 次の段落を追加 → …」という逐次処理だった。

```csharp
// NG パターン（旧実装）
dynamic para = doc.Paragraphs.Add();
para.Range.Text = text;
para.Range.Style = headingStyle;   // ← ここで Word 内部の再計算がペンディングになる

para = doc.Paragraphs.Add();       // ← ペンディング中に呼ぶと参照がずれる
```

Word COM は `Range.Style` への代入後に内部の段落再計算を **非同期的に** キューに積む。
このペンディング中に `Paragraphs.Add()` を呼ぶと、返ってくる参照が
期待の段落ではなく別の段落を指す。結果として後続の `Range.Text` 代入が
1 段落前を上書きし、末尾段落が失われたように見える。

`Paragraphs.Count` を都度取得して末尾参照に切り替える方法も試みたが、
`ApplyBulletDefault()` が内部で継続段落を暗黙追加するため Count がずれ、
同様の参照ミスが交互に発生した。

---

## 解決策: 4 フェーズ方式（v0.3.0 採用）

「段落を作りながらスタイルを当てる」操作を完全に分離する。

### Phase 1 — ParseBlocks（COM 不使用）

Markdown を `List<Block>` に変換する純粋関数。
テーブル・見出し・箇条書き・HR・改ページ・空行・通常段落を判別する。
見出し番号付与もここで文字列として付与する（Word 自動アウトラインは不使用）。

```csharp
private enum BlockKind { Empty, Paragraph, Heading, Bullet, Hr, PageBreak, Table }
```

### Phase 2 — WriteAllText（1 回の COM 呼び出し）

全段落テキストを `\r` 区切りで結合し `doc.Content.Text` に一括代入する。

```csharp
doc.Content.Text = "見出し1\r箇条書き1\r箇条書き2\r...";
```

この 1 回の代入で Word 内部の段落構造が確定し、
以降は `doc.Paragraphs[i + 1]` による安定したインデックスアクセスが可能になる。

改ページは `\f`（Chr(12) = wdPageBreak）をテキストとして埋め込む。

### Phase 3 — ApplyStyles（インデックス基準、2 パス）

確定済みの段落構造に対してスタイルを適用する。
`Paragraphs.Add()` は一切呼ばない。

**パス 1**: 各ブロックに段落スタイル・フォントを設定する。

| BlockKind | 処理 |
|-----------|------|
| Heading   | `para.Range.Style = -(level + 1)` (WdBuiltinStyle) |
| Bullet    | `para.Range.Style = -47` (wdStyleListBullet) |
| Paragraph | `para.Range.Style = -1` (wdStyleNormal) + フォント |
| Empty     | `para.Range.Style = -1` |
| Hr        | `para.Range.Style = -1` + `Borders[-3].LineStyle = 1` |
| PageBreak | なし（\f で確定済み） |

**パス 2**: 非箇条書き段落から自動伝播リスト書式を除去する。

```csharp
if ((int)para.Range.ListFormat.ListType != 0) // 0 = wdListNoNumbering
    para.Range.ListFormat.RemoveNumbers();
```

2 パスが必要な理由: Word は `wdStyleListBullet` 適用時に隣接段落へ
リスト書式を **直接適用形式** で伝播する。これはスタイル設定（`para.Range.Style = -1`）
だけでは除去できない。また順方向ループでは `Bullet[i]` 適用後に `Para[i-1]`
へ逆方向伝播した書式を拾えないため、パス 1 完了後にまとめて除去する必要がある。

### Phase 4 — InsertTables（逆順）

テーブルはプレースホルダー段落の位置に `doc.Tables.Add(para.Range, rows, cols)` で挿入する。
挿入によって後続の段落インデックスがシフトするため、末尾から逆順に処理する。

---

## 重要な Word COM の挙動・定数メモ

### 段落区切り文字

| 文字 | 定数 | 意味 |
|------|------|------|
| `\r` | — | 段落マーク（Paragraph Mark） |
| `\v` | Chr(11) | ソフトリターン（段落内改行） |
| `\f` | Chr(12) | 改ページ（wdPageBreak） |

### WdBuiltinStyle（言語非依存の定数値）

| 値 | スタイル |
|----|----------|
| -1 | wdStyleNormal（標準） |
| -2 | wdStyleHeading1（見出し 1） |
| -3 | wdStyleHeading2（見出し 2） |
| -4 | wdStyleHeading3（見出し 3） |
| -47 | wdStyleListBullet（箇条書き） |

### ListType 値（ListFormat.ListType）

| 値 | 意味 |
|----|------|
| 0 | wdListNoNumbering（リストなし） |
| 2 | wdListBullet |
| 3 | wdListSimpleNumbering |

### `ApplyBulletDefault()` の罠

このメソッドは **トグル動作** する。
段落にすでにリスト書式が存在する場合、呼び出すと除去される。
Word が隣接段落への自動伝播でリスト書式を付けた後に呼ぶと消失する。
→ `para.Range.Style = -47` に置き換えること。

### COM オブジェクト解放

```csharp
try { doc.Close(false); } catch { }
Marshal.FinalReleaseComObject(doc);
try { app.Quit(false); } catch { }
Marshal.FinalReleaseComObject(app);
```

`Close` / `Quit` は例外が飛んでも `FinalReleaseComObject` を必ず実行する。

---

## v0.3.0 の既知の残課題

- 箇条書き記号の**見た目**が Word テンプレートのデフォルト（wdStyleListBullet）に依存する。
  カスタムテンプレートを使うユーザーでは記号が異なる可能性がある。
- 順序付きリスト（`1.` / `2.`）は現在 `BlockKind.Paragraph` として処理される（箇条書き非対応）。
  将来対応する場合は `wdStyleListNumber`（-48）を検討する。
- 見出し番号の書式（`1.` / `1.1` / `1.1.1`）は固定。区切り文字やゼロ埋め等は未対応。
- 見出しスタイルのフォント・色は Word テンプレート定義に完全依存（フォント上書き機能は v0.3.0 で廃止）。

---

## 試みて失敗した方法（再実施を避けるために記録）

| 方法 | 問題点 |
|------|--------|
| `Paragraphs.Add()` → Style → `Paragraphs.Add()` … の逐次処理 | Word 内部の非同期再計算により後続の参照がずれ段落消失 |
| `doc.Paragraphs[doc.Paragraphs.Count]` で末尾参照 | `ApplyBulletDefault` が継続段落を暗黙追加し Count がずれる |
| `ApplyBulletDefault()` を Style 適用前に移動 | 継続段落への自動伝播が起き 2 番目の箇条書きで記号消失 |
| Style 適用直後に `_ = para.Range.Text` を読み取り（強制同期試行） | COM の再計算タイミングが変わらず悪化（セクション末尾のみ残る） |
| `para.Range.Style = -1` のみで伝播リスト書式をクリア | Style 設定は段落スタイルを変えるだけで直接適用 ListFormat は除去されない |
