# YouTube2Slides

YouTube動画から簡単にGoogleスライドを作成するツール。特に3Dモデリングなどのチュートリアル動画を効率的に復習するために開発されました。

## 新機能

- **字幕量に基づく自動分割**: 一定量の字幕テキストが溜まったら自動的に新しいスライドを作成
- **画面変化検出**: 動画内で大きな画面変化があった時に自動的にスライドを分割
- **従来の等間隔抽出との併用**: 上記の機能を従来の時間間隔ベースの抽出と併用可能
- **高品質画像出力**: 画像品質を最大100%まで調整可能
- **動画品質設定**: 様々な解像度に対応する動画品質設定
- **Whisper API連携**: OpenAIのWhisper APIを使用した高精度な音声認識に対応（NEW!）

## 特徴

- YouTubeの動画から一定間隔でスクリーンショットを自動抽出
- 字幕テキストを自動取得してスライドに追加
- タイムスタンプ付きの目次を自動生成
- 各スライドに元動画へのリンク（タイムスタンプ付き）
- GoogleSlidesで開ける形式で保存
- 画像を最適なサイズで表示
- テキストの改行が少ないコンパクト表示
- Whisper APIによる高精度な音声認識（NEW!）

## インストール方法

### 前提条件
- Python 3.7以上
- pip (Pythonパッケージマネージャー)
- FFmpeg (音声処理に必要、Whisper API使用時)

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
pip install yt-dlp youtube-transcript-api opencv-python pillow python-pptx requests
```

### FFmpegのインストール（Whisper API使用時に必要）

#### Windows:
1. [FFmpeg公式サイト](https://ffmpeg.org/download.html)からダウンロード
2. 解凍してbinフォルダを環境変数PATHに追加

#### macOS:
```
brew install ffmpeg
```

#### Linux:
```
sudo apt update && sudo apt install ffmpeg  # Debian/Ubuntu
sudo yum install ffmpeg  # CentOS/RHEL
```

## 使用方法

### コマンドライン

基本的な使い方：
```
python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID"
```

利用可能なオプション：
```
usage: youtube_tutorial_slides.py [-h] [--output OUTPUT] [--format {pptx,slides}] [--interval INTERVAL] [--lang LANG] [--nocompact] [--text-threshold TEXT_THRESHOLD] [--change-threshold CHANGE_THRESHOLD] [--disable-text-split] [--disable-change-detection] [--use-whisper] [--whisper-api-key WHISPER_API_KEY] [--whisper-model {tiny,base,small,medium,large-v1,large-v2,large-v3}] [--no-force-whisper] url

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
  --use-whisper                 Whisper APIを使用して高精度な音声認識を行う（デフォルトで字幕優先使用）
  --whisper-api-key WHISPER_API_KEY
                                Whisper API（OpenAI API）のキー。環境変数 OPENAI_API_KEY からも取得可能
  --whisper-model {tiny,base,small,medium,large-v1,large-v2,large-v3}
                                使用するWhisperモデル（注: 現在のOpenAI APIではすべて内部的に同じモデルを使用）
  --no-force-whisper            Whisper API使用時に品質比較を行い、必ずしもWhisper APIの結果を優先しない
```

### バッチファイル（Windows）

Windowsユーザーの場合、同梱の `enhanced_run_slides_tool.bat` をダブルクリックするだけで実行できます。画面の指示に従って必要な情報を入力してください。

## 使用例

1. 60秒間隔でスクリーンショットを取得する場合：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --interval 60
   ```

2. 英語の字幕を使用する場合：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --lang en
   ```

3. 字幕量300文字ごとにスライドを分割する場合：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --text-threshold 300
   ```

4. 画面変化検出の感度を高くする場合（0.4に設定）：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --change-threshold 0.4
   ```

5. 画面変化検出のみを使用する場合（文字量による分割を無効化）：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --disable-text-split
   ```

6. 画像品質を最高に設定する場合：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --image-quality 100
   ```

7. 動画の最高品質を使用する場合：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --video-quality best
   ```

8. Whisper APIを使用して高精度な音声認識を行う場合（デフォルトでWhisper結果優先）：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --use-whisper --whisper-api-key "YOUR_API_KEY"
   ```

9. Whisper APIを使用するが、品質比較を行う場合：
   ```
   python youtube_tutorial_slides.py "https://www.youtube.com/watch?v=VIDEO_ID" --use-whisper --whisper-api-key "YOUR_API_KEY" --no-force-whisper
   ```

## GoogleSlidesでの使用方法

1. [Google Drive](https://drive.google.com/)にアクセスしてログイン
2. 「新規」→「ファイルをアップロード」から生成された `_googleslides.pptx` ファイルをアップロード
3. アップロードしたファイルを右クリックして「アプリで開く」→「Googleスライド」を選択
4. Googleスライドが開いたら「ファイル」→「名前を付けて保存」でGoogleスライド形式として保存

## Whisper API 使用時の注意点

1. **APIキーの入手**: OpenAIのアカウントを作成し、APIキーを取得してください。
2. **コスト管理**: Whisper APIは使用量に応じて課金されます。特に長い動画では注意が必要です。
3. **モデル選択**: 現在のOpenAI APIでは、APIリクエスト時にはすべてのモデルサイズが内部的に同じエンドポイントを使用します。
4. **字幕優先使用**: `--use-whisper` オプションを使用すると、デフォルトで常にWhisper APIの結果を優先使用します。
5. **品質比較モード**: `--no-force-whisper` オプションを追加すると、YouTube字幕とWhisper字幕の品質を比較して良い方を選択します。

## トラブルシューティング

- **字幕が取得できない**：動画に字幕がない可能性があります。別の言語を試すか、自動生成字幕の有無を確認してください。
- **動画のダウンロードに失敗**：インターネット接続を確認するか、別のYouTube動画URLで試してください。
- **スライドのレイアウトが崩れる**：GoogleSlidesでのインポート後、手動で調整が必要な場合があります。
- **画面変化検出がうまく機能しない**：`--change-threshold` の値を調整してみてください。値を小さくするとより敏感になります。
- **スライドが多すぎる**：文字量閾値を大きくするか、`--disable-text-split` や `--disable-change-detection` を使って一部の分割機能を無効にしてみてください。
- **Whisper APIエラー**：APIキーが正しいか確認してください。また、環境変数 `OPENAI_API_KEY` を設定することでもAPIキーを提供できます。
- **FFmpegエラー**：Whisper API使用時にはFFmpegが必要です。インストールされていることを確認してください。

## パラメータ調整のヒント

- **text-threshold**: 一般的なチュートリアル動画では200〜300文字程度が適切です。速いペースで解説が進む動画では値を大きく、ゆっくり解説している動画では小さくすると良いでしょう。
- **change-threshold**: デフォルト値の0.6は中程度の感度です。3Dモデリングなど視覚的な変化が多い動画では0.7〜0.8に、画面の変化が少ない講義形式の動画では0.4〜0.5に調整するのがおすすめです。
- **whisper-model**: 現在のOpenAI APIではモデルサイズによる違いはありませんが、UIの互換性のために選択肢を提供しています。

## 開発者向け情報

このプロジェクトは以下のライブラリを使用しています：

- **yt-dlp**: YouTubeからの動画ダウンロード
- **youtube-transcript-api**: YouTubeの字幕取得
- **opencv-python**: 動画からのフレーム抽出と画像解析
- **pillow**: 画像処理
- **python-pptx**: PowerPoint形式のファイル生成
- **requests**: Whisper APIとの通信

## ライセンス

このプロジェクトは [MIT License](LICENSE) の下で公開されています。

## 免責事項

このツールは教育目的でのみ使用してください。YouTubeの利用規約に従ってコンテンツを取り扱い、著作権を尊重してください。
