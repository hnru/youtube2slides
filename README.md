# YouTube2Slides

YouTube動画から簡単にGoogleスライドを作成するツール。特に3Dモデリングなどのチュートリアル動画を効率的に復習するために開発されました。

<p align="center">
  <img src="https://github.com/yourusername/youtube2slides/raw/main/docs/sample_slide.png" alt="サンプルスライド" width="600">
</p>

## 特徴

- YouTubeの動画から一定間隔でスクリーンショットを自動抽出
- 字幕テキストを自動取得してスライドに追加
- タイムスタンプ付きの目次を自動生成
- 各スライドに元動画へのリンク（タイムスタンプ付き）
- GoogleSlidesで開ける形式で保存
- 画像を最適なサイズで表示
- テキストの改行が少ないコンパクト表示

## インストール方法

### 前提条件
- Python 3.7以上
- pip (Pythonパッケージマネージャー)

### インストール手順

1. リポジトリをクローンまたはダウンロードします。
   ```
   git clone https://github.com/yourusername/youtube2slides.git
   cd youtube2slides
   ```

2. 必要なパッケージをインストールします。
   ```
   pip install -r requirements.txt
   ```

または直接必要なパッケージをインストールすることもできます：
```
pip install yt-dlp youtube-transcript-api opencv-python pillow python-pptx
```

## 使用方法

### コマンドライン

基本的な使い方：
```
python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID"
```

利用可能なオプション：
```
usage: youtube_tutorial_slides.py [-h] [--output OUTPUT] [--interval INTERVAL] [--lang LANG] [--nocompact] url

YouTube チュートリアル動画からGoogleSlides用スライドを生成

positional arguments:
  url                    YouTube 動画の URL

optional arguments:
  -h, --help             このヘルプメッセージを表示して終了
  --output OUTPUT, -o OUTPUT
                         出力ディレクトリ (デフォルト: "output")
  --interval INTERVAL, -i INTERVAL
                         フレーム抽出間隔（秒） (デフォルト: 30)
  --lang LANG, -l LANG   字幕の言語コード (デフォルト: "ja")
  --nocompact, -nc       テキストを改行せずにまとめない
```

### バッチファイル（Windows）

Windowsユーザーの場合、同梱の `run_slides_tool.bat` をダブルクリックするだけで実行できます。画面の指示に従って必要な情報を入力してください。

## 使用例

1. 60秒間隔でスクリーンショットを取得する場合：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --interval 60
   ```

2. 英語の字幕を使用する場合：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --lang en
   ```

3. 出力ディレクトリを指定する場合：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --output "my_tutorials"
   ```

## GoogleSlidesでの使用方法

1. [Google Drive](https://drive.google.com/)にアクセスしてログイン
2. 「新規」→「ファイルをアップロード」から生成された `_googleslides.pptx` ファイルをアップロード
3. アップロードしたファイルを右クリックして「アプリで開く」→「Googleスライド」を選択
4. Googleスライドが開いたら「ファイル」→「名前を付けて保存」でGoogleスライド形式として保存

## トラブルシューティング

- **字幕が取得できない**：動画に字幕がない可能性があります。別の言語を試すか、自動生成字幕の有無を確認してください。
- **動画のダウンロードに失敗**：インターネット接続を確認するか、別のYouTube動画URLで試してください。
- **スライドのレイアウトが崩れる**：GoogleSlidesでのインポート後、手動で調整が必要な場合があります。

## 開発者向け情報

このプロジェクトは以下のライブラリを使用しています：

- **yt-dlp**: YouTubeからの動画ダウンロード
- **youtube-transcript-api**: YouTubeの字幕取得
- **opencv-python**: 動画からのフレーム抽出
- **pillow**: 画像処理
- **python-pptx**: PowerPoint形式のファイル生成

## ライセンス

このプロジェクトは [MIT License](LICENSE) の下で公開されています。

## 免責事項

このツールは教育目的でのみ使用してください。YouTubeの利用規約に従ってコンテンツを取り扱い、著作権を尊重してください。
