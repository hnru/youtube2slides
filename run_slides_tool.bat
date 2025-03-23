@echo off
echo YouTube動画からGoogleスライド作成ツール（拡張版）
echo ================================================
echo.
echo 必要なライブラリをインストールしています...
pip install yt-dlp youtube-transcript-api opencv-python pillow python-pptx
echo.
echo ダウンロードしたい YouTube 動画の URL を入力してください:
set /p youtube_url=URL: 
echo.
echo フレーム抽出間隔（秒）を入力してください (デフォルト: 30):
set /p interval=間隔: 
echo.
echo 字幕の言語コードを入力してください (デフォルト: ja):
set /p lang=言語コード: 
echo.
echo テキストをコンパクト表示しますか? (Y/N) (デフォルト: Y):
set /p compact_choice=選択: 

if "%interval%"=="" set interval=30
if "%lang%"=="" set lang=ja
if "%compact_choice%"=="" set compact_choice=Y

set compact_param=
if /i "%compact_choice%"=="N" set compact_param=--nocompact

echo.
echo --- 拡張機能設定 ---
echo 文字量によるスライド分割を有効にしますか? (Y/N) (デフォルト: Y):
set /p text_split_choice=選択: 
echo.
echo スライド分割を行う文字数閾値を入力してください (デフォルト: 200):
set /p text_threshold=閾値: 
echo.
echo 画面変化検出によるスライド分割を有効にしますか? (Y/N) (デフォルト: Y):
set /p change_detect_choice=選択: 
echo.
echo 画像品質を設定してください (1-100, 高いほど高品質。デフォルト: 95):
set /p image_quality=画像品質: 
echo.
echo 動画ダウンロード品質を選択してください:
echo 1: 最高品質 (best)
echo 2: 高品質 - 1080p (high)
echo 3: 中品質 - 720p (medium)
echo 4: 低品質 - 480p (low)
set /p video_quality_choice=選択 (1-4): 

if "%image_quality%"=="" set image_quality=95

set video_quality=high
if "%video_quality_choice%"=="1" set video_quality=best
if "%video_quality_choice%"=="2" set video_quality=high
if "%video_quality_choice%"=="3" set video_quality=medium
if "%video_quality_choice%"=="4" set video_quality=low

set text_split_param=
if /i "%text_split_choice%"=="N" set text_split_param=--disable-text-split

set change_detect_param=
if /i "%change_detect_choice%"=="N" set change_detect_param=--disable-change-detection

echo.
echo 処理を開始します...
python enhanced_youtube_slides.py "%youtube_url%" --interval %interval% --lang %lang% %compact_param% --text-threshold %text_threshold% %text_split_param% --change-threshold %change_threshold% %change_detect_param% --image-quality %image_quality% --video-quality %video_quality%
echo.
echo 処理が完了しました！
echo 任意のキーを押して終了...
pause > nul