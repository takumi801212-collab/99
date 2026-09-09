# デカ盛りチャレンジ 提案スライド ジェネレータ

`docs/dekamori-tagline.pptx`（全15枚・16:9）を生成するスクリプトです。
スライドは手で編集せず、**必ず `gen.js` を編集して再生成**してください。

## 再生成

```bash
cd tools/pptx
npm install                 # pptxgenjs
node gen.js                 # 出力先は gen.js 末尾の writeFile() で指定
```

出力パスは `gen.js` 最終行にべた書きしています。環境に合わせて
`/path/to/repo/docs/dekamori-tagline.pptx` へ書き換えてください。

## 検証（任意）

```bash
pip install defusedxml lxml Pillow "markitdown[pptx]"
python3 <pptx-skill>/scripts/office/validate.py ../../docs/dekamori-tagline.pptx
```

## 目視チェック用のレンダリング

LibreOffice の Impress フィルタが必要です。`libreoffice-core` だけでは
`.pptx` を読めず `Error: source file could not be loaded` になります。

```bash
apt-get install -y libreoffice-impress
soffice --headless --convert-to pdf --outdir . ../../docs/dekamori-tagline.pptx
# pdftoppm が無い環境では pypdfium2 で代替
python3 -c "
import pypdfium2 as p
d=p.PdfDocument('dekamori-tagline.pdf')
for i in range(len(d)):
    d[i].render(scale=110/72).to_pil().convert('RGB').save(f'slide-{i+1:02d}.jpg',quality=85)"
```

## デザイン仕様

| 項目 | 値 |
|---|---|
| 版面 | LAYOUT_WIDE（13.333 × 7.5 inch） |
| 和文書体 | Meiryo（Windows/Mac の PowerPoint に同梱。Aptos は使わない） |
| 墨 / 紙 | `16181C` / `FFFFFF` |
| 面 | `F1F1EC` / `E7E7E0` |
| 企画色（POP赤） | `D8232A` |
| 強調（蛍光イエロー） | `F7DE3A` |
| 副文字 | `6E7076`（明地）/ `9BA0A8`（暗地） |
| モチーフ | 店頭プライスカード＝赤地のチップ。全編で反復 |
| 構成 | 表紙・結論・骨子・次の一手をダークに。中身はライト（サンドイッチ） |

### 注意点（踏んだ地雷）

- `chip()` の幅は `0.34 + 文字数 × 0.158`。これ未満だと和文が枠から溢れる
- 色コードに `#` を付けると **ファイルが壊れる**
- オプションオブジェクトは呼び出しごとに新規生成する（pptxgenjs が破壊的に書き換える）
- 各スライドの最下部テキストは、直上のカード下端と 0.2 inch 以上あけること
