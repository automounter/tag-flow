# Tag FloW

MP3 のタグ（ID3）を、STEP を順にたどって整えるデスクトップアプリです。Python と tkinter で動き、macOS・Windows・Linux で使えます。

*A desktop app (Python / tkinter) that cleans up MP3 ID3 tags step by step: extract tags to Excel, convert Japanese titles to English, decide album artists and genres, check artwork, write the tags back, and rename / sort the files. The UI is in Japanese.*

> [!WARNING]
> **このアプリは MP3 ファイルのタグを書き換え、ファイルの名前を変えたり移したりします。使う前に、必ず MP3 のバックアップを取ってください。** 最初は、コピーした少しの曲で試すのがおすすめです。

## できること

| STEP | 内容 |
|---|---|
| 01 | 小さすぎるファイル（Mac の補足ファイル・壊れたファイルなど）を探し、選んで削除 |
| 02 | フォルダの中の MP3 のタグを読み、Excel に書き出す（以降の STEP の作業データ） |
| 03 | 邦題・日本語のアーティスト名を、辞書と Web の照会（Apple Music・Discogs・Spotify）で英語表記に変換 |
| 04 | アルバムごとに、アルバムアーティスト（単独・複数並記・Various Artists）を決める |
| 05 | アルバムごとに、Discogs のジャンル＞スタイルを決める |
| 06 | アルバム写真を Apple Music・Spotify・Discogs と照らし合わせる（任意） |
| 07 | 変更点を確かめて、MP3 のタグに書き込む（書き込む前のタグは控えを取り、戻せる） |
| 08 | ファイル名を曲名に変え、アーティスト別のフォルダへコピーする（元のファイルの削除も選べる） |

- **タグ一覧編集**：Excel のような表で、タグを直接直す（コピー＆ペースト・一括入力・元に戻す・検索置換・ファイル名とタグの相互変換）
- **サイトから取り寄せる**：MP3 データベースのサイトに登録されたタグを、手元の MP3 に書く（サイトの鍵が必要）

## 動く環境

試した OS：macOS、Windows 11、Ubuntu 26.04、Debian 13、AlmaLinux 9（RHEL 系）

- Python 3.9 以上（tkinter が入っていること）
- 部品：mutagen・openpyxl・Pillow・tksheet（`requirements.txt`）

## 入れ方

### 1. Python を用意する

| OS | 方法 |
|---|---|
| macOS | [python.org](https://www.python.org/downloads/) の公式の Python を入れる（tkinter が付いている） |
| Windows | [python.org](https://www.python.org/downloads/) の公式の Python を入れる。入れるときに「Add python.exe to PATH」に印を付ける |
| Ubuntu・Debian | `sudo apt install python3 python3-tk python3-venv` |
| RHEL・AlmaLinux・Rocky | `sudo dnf install python3 python3-tkinter` |

Linux で日本語の字体が無いときは、Ubuntu・Debian は `fonts-noto-cjk`、RHEL 系は `google-noto-sans-cjk-ttc-fonts` も入れてください。

### 2. Tag FloW を取ってきて、部品を入れる

```sh
git clone https://github.com/automounter/tag-flow.git
cd tag-flow
python3 -m venv .venv
```

続けて、OS ごとに次を実行します。

```sh
# macOS・Linux
.venv/bin/pip install -r requirements.txt
```

```bat
rem Windows
.venv\Scripts\pip install -r requirements.txt
```

### 3. 起動する

```sh
# macOS・Linux
.venv/bin/python main.py
```

```bat
rem Windows
.venv\Scripts\python main.py
```

macOS では、Finder で `run.command` をダブルクリックしても起動できます。

## 最初の設定

右上の「≡」から設定を開き、次を指定します。

- **MP3 保存元フォルダ**：整える MP3 が入っているフォルダ
- **出力先フォルダ**：STEP 08 でコピーする先
- **Excel 保存先フォルダ**：STEP 02 で書き出す Excel の置き場所
- **辞書データ フォルダ**：英語表記・ジャンルの辞書の置き場所。同梱の `data/` を使うときは空のままでよい

設定は次の場所に保存されます。

| OS | 場所 |
|---|---|
| macOS | `~/Library/Application Support/TagFlow/` |
| Windows | `%APPDATA%\TagFlow\` |
| Linux | `~/.config/TagFlow/` |

### Discogs・Spotify のキー（任意）

STEP 03 の名前照会・STEP 05・STEP 06 で Discogs と Spotify も使うときは、キーを書いたテキストファイルを作り、設定の「照会サービスのキー」で指定します。キーが無くても、Apple Music だけで動きます。

```text
【Discogs】
token: あなたのDiscogsのトークン

【Spotify】
クライアントID: あなたのクライアントID
クライアントシークレット: あなたのクライアントシークレット
```

- Discogs のトークン：<https://www.discogs.com/settings/developers>
- Spotify のクライアント ID：<https://developer.spotify.com/dashboard>

キーのファイルは、このフォルダの外に置いてください。アプリはキーをどこにも書き出しません。

### AI で補う（任意）

STEP 03 の名前照会で、見つからなかったものを AI に補ってもらう機能があります。[Claude Code](https://claude.com/claude-code) の `claude` か、[Codex CLI](https://github.com/openai/codex) の `codex` が、その PC で使える（ログイン済みの）ときだけ動きます。

## 同梱の辞書

`data/` に、英語表記とジャンルの辞書（Excel）が入っています。作者の手元の曲から作ったものなので、足りない分は STEP 03・05 の画面や Excel で足してください。

- `translation_dictionary.xlsx`・`translation_overrides.xlsx`：日本語 → 英語表記
- `genre_styles.xlsx`・`genre_dictionary.xlsx`：ジャンル・スタイルと、アーティストの既定のジャンル
- `artist_groups.txt`：「&」「,」を含んでも分けないグループ名

## 知られている問題

- Linux・Windows では、字体の違いで、画面の文字の並びが macOS と少し違います
- 実行ファイル（.exe・.app）は配っていません。署名の無い実行ファイルは、Windows の Smart App Control や macOS の Gatekeeper に止められるためです。Python から起動してください

## 不具合の報告

[Issues](https://github.com/automounter/tag-flow/issues) から知らせてください。OS とその版、Python の版、どの STEP で何をしたか、出たメッセージを書いてもらえると助かります。ログは設定の場所の `logs/` にあります（曲名やフォルダ名が入るので、貼る前に確かめてください）。

## ライセンス

[MIT License](LICENSE)
