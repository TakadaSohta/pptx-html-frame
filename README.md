# HTML Frame

PowerPoint のスライドに HTML や Web ページを置き、スライドショー中も操作できるようにするアドインです（Office のコンテンツ アドイン）。

アドインの画面は GitHub Pages（https://takadasohta.github.io/pptx-html-frame/）から読み込むので、サーバーの起動や証明書の準備は要りません。

## できること

- **HTML を書く**: HTML を書くか貼り付けると、スライドの枠の中で動きます。HTML ファイルの読み込みもできます。
- **URL を開く**: https のページを枠の中に表示します。
- **基準の幅**: 1280 などを入れると、その幅で描いた画面を枠に合わせて拡大縮小します。編集画面とスライドショーでは枠のピクセル数が変わるので、見た目をそろえたいときに使います。
- 設定は枠ごとにプレゼンテーション（.pptx）の中へ保存されます。1 枚のスライドに複数置けます。書いた HTML は GitHub には送られません。

## PC に登録する（最初の一度だけ）

Node.js が入った Windows の PC で、このリポジトリを取ってきて次を実行します。

```
npm install
npm run try
```

`npm run try` はアドインを登録し、HTML Frame を置いたプレゼンテーションを PowerPoint で開きます。ほかのスライドやプレゼンテーションで使うときは、この枠をコピーして貼り付けます。貼り付けた枠は、それぞれ別の内容を持てます。

登録だけするときは `npm run register` を使います。PowerPoint の「アプリ」画面（ストア）には HTML Frame は出ません。

Node.js がない PC では、PowerShell で次を実行しても同じです（`<manifest.xml の場所>` は manifest.xml を保存したフルパス）。

```
New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer" -Force
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer" -Name "d7a29fd6-c7ec-4e91-aa67-e9044110eb30" -Value "<manifest.xml の場所>" -PropertyType String -Force
```

登録を外すときは `npm run unregister` を実行します。

## 使い方

- 枠の中の編集画面にサンプルが入っています。「保存して表示」を押すと動きます。
- 枠の右上の「編集」で内容を変え、「↻」で最初から表示し直します。スライドショー中はこのボタンは出ません。

## 注意

- アドインの画面と Office の部品をインターネットから読み込むので、オフラインの会場では動きません。
- スライドショー中に枠の中をクリックすると、キー入力が枠の中に入ります。矢印キーでスライドが進まないときは、枠の外をクリックします。
- 埋め込みを禁止しているサイト（Google 検索、GitHub など）は URL を開くモードでも表示されません。YouTube は `https://www.youtube.com/embed/動画ID` の形で入れます。
- 読み込んだ HTML ファイルから相対パスで参照している画像や JS は読み込まれません。1 つのファイルにまとめるか、CDN の URL で読み込みます。
- アドインを登録していない PC で .pptx を開くと、枠には最後の状態の静止画が出ます。
- `docs/` を書き換えて main に push すると、GitHub Pages を通じて全部の PC のアドインが新しくなります。壊れた状態で push しないよう、先に開発版で確かめます。

## 開発

`docs/` を手元で書き換えて試すときは、localhost を読む開発版（manifest.dev.xml、表示名「HTML Frame（開発版）」）を使います。

```
npm run certs
npm start
```

1. `npm run certs` で localhost 用の開発証明書を入れます。Windows が証明書を信頼するか尋ねるので「はい」を選びます。
2. `npm start` で `https://localhost:3000` のサーバーが起動し、開発版を置いたテスト用のプレゼンテーションが PowerPoint で開きます。
3. 2 回目からは `npm run dev-server` でサーバーだけ起動すれば使えます。`npm run stop` はサーバーを止め、開発版の登録も外します。

Office の外のブラウザで確かめるときは `npm run serve:http` を実行して `http://localhost:3000/` を開きます。設定はブラウザに保存され、`?view=read` を付けるとスライドショー中の表示になります。

## ファイル

| パス | 中身 |
| --- | --- |
| `manifest.xml` | アドインの定義（GitHub Pages を読む） |
| `manifest.dev.xml` | 開発版の定義（https://localhost:3000 を読む） |
| `docs/index.html`, `docs/app.js`, `docs/style.css` | 枠の中に出る画面。GitHub Pages でそのまま公開される |
| `docs/sample.html` | 最初に入るサンプル |
| `serve.js` | `docs/` を配信する開発用サーバー |
