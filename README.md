# YouTube2Slides

YouTube動画から簡単にGoogleスライドを作成するツール。特に3Dモデリングなどのチュートリアル動画を効率的に復習するために開発されました。

## 新機能

- **字幕量に基づく自動分割**: 一定量の字幕テキストが溜まったら自動的に新しいスライドを作成
- **画面変化検出**: 動画内で大きな画面変化があった時に自動的にスライドを分割
- **従来の等間隔抽出との併用**: 上記の機能を従来の時間間隔ベースの抽出と併用可能
- **高品質画像出力**: 画像品質を最大100%まで調整可能
- **動画品質設定**: 様々な解像度に対応する動画品質設定

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
python enhanced_youtube_slides.py "https://www.youtube.com/watch?v=VIDEO_ID"
```

利用可能なオプション：
```
usage: enhanced_youtube_slides.py [-h] [--output OUTPUT] [--format {pptx,slides}] [--interval INTERVAL] [--lang LANG] [--nocompact] [--text-threshold TEXT_THRESHOLD] [--change-threshold CHANGE_THRESHOLD] [--disable-text-split] [--disable-change-detection] url

YouTube チュートリアル動画からGoogleSlides用スライドを生成（拡張版）

positional arguments:
  url                            YouTube 動画の URL

optional arguments:
  -h, --help                     このヘルプメッセージを表示して終了
  --output OUTPUT, -o OUTPUT     出力ディレクトリ (デフォルト: "output")
  --format {pptx,slides}, -f {pptx,slides}
                                出力形式 (デフォルト: "slides")
  --interval INTERVAL, -i INTERVAL
                                フレーム抽出間隔（秒） (デフォルト: 30)
  --lang LANG, -l LANG          字幕の言語コード (デフォルト: "ja")
  --nocompact, -nc              テキストを改行せずにまとめない
  
  --text-threshold TEXT_THRESHOLD, -tt TEXT_THRESHOLD
                                スライド分割を行う文字数閾値 (デフォルト: 200文字)
  --change-threshold CHANGE_THRESHOLD, -ct CHANGE_THRESHOLD
                                画面変化検出の閾値 (0-1, 低いほど敏感。デフォルト: 0.6)
  --disable-text-split, -dts    文字量によるスライド分割を無効にする
  --disable-change-detection, -dcd
                                画面変化検出によるスライド分割を無効にする
  --image-quality IMAGE_QUALITY, -iq IMAGE_QUALITY
                                画像品質 (1-100, 高いほど高品質。デフォルト: 95)
  --video-quality {best,high,medium,low}, -vq {best,high,medium,low}
                                動画ダウンロード品質 (best=最高品質, high=高品質1080p, 
                                medium=中品質720p, low=低品質480p。デフォルト: high)
```

### バッチファイル（Windows）

Windowsユーザーの場合、同梱の `enhanced_run_slides_tool.bat` をダブルクリックするだけで実行できます。画面の指示に従って必要な情報を入力してください。

## 使用例

1. 60秒間隔でスクリーンショットを取得する場合：
   ```
   python enhanced_youtube_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --interval 60
   ```

2. 英語の字幕を使用する場合：
   ```
   python enhanced_youtube_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --lang en
   ```

3. 字幕量300文字ごとにスライドを分割する場合：
   ```
   python enhanced_youtube_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --text-threshold 300
   ```

4. 画面変化検出の感度を高くする場合（0.4に設定）：
   ```
   python enhanced_youtube_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --change-threshold 0.4
   ```

5. 画面変化検出のみを使用する場合（文字量による分割を無効化）：
   ```
   python enhanced_youtube_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --disable-text-split
   ```

6. 画像品質を最高に設定する場合：
   ```
   python enhanced_youtube_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --image-quality 100
   ```

7. 動画の最高品質を使用する場合：
   ```
   python enhanced_youtube_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --video-quality best
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
- **画面変化検出がうまく機能しない**：`--change-threshold` の値を調整してみてください。値を小さくするとより敏感になります。
- **スライドが多すぎる**：文字量閾値を大きくするか、`--disable-text-split` や `--disable-change-detection` を使って一部の分割機能を無効にしてみてください。

## パラメータ調整のヒント

- **text-threshold**: 一般的なチュートリアル動画では200〜300文字程度が適切です。速いペースで解説が進む動画では値を大きく、ゆっくり解説している動画では小さくすると良いでしょう。
- **change-threshold**: デフォルト値の0.6は中程度の感度です。3Dモデリングなど視覚的な変化が多い動画では0.7〜0.8に、画面の変化が少ない講義形式の動画では0.4〜0.5に調整するのがおすすめです。

## 開発者向け情報

このプロジェクトは以下のライブラリを使用しています：

- **yt-dlp**: YouTubeからの動画ダウンロード
- **youtube-transcript-api**: YouTubeの字幕取得
- **opencv-python**: 動画からのフレーム抽出と画像解析
- **pillow**: 画像処理
- **python-pptx**: PowerPoint形式のファイル生成

## ライセンス

このプロジェクトは [MIT License](LICENSE) の下で公開されています。

## 免責事項

このツールは教育目的でのみ使用してください。YouTubeの利用規約に従ってコンテンツを取り扱い、著作権を尊重してください。
