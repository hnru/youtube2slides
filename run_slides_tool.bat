@echo off
echo YouTube動画からGoogleスライド作成ツール（拡張版 with Whisper API）
echo ================================================================
echo.
echo 必要なライブラリをインストールしています...
pip install yt-dlp youtube-transcript-api opencv-python pillow python-pptx requests
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
echo.
echo --- Whisper API 設定 ---
echo 高精度な音声認識のためにWhisper APIを使用しますか? (Y/N) (デフォルト: N):
set /p whisper_choice=選択: 
echo.

set whisper_param=
set no_force_whisper_param=
if /i "%whisper_choice%"=="Y" (
    echo OpenAI API Key を入力してください:
    set /p whisper_api_key=API Key: 
    echo.
    echo Whisperモデルを選択してください (注: 現在のOpenAI APIではすべて内部的に同じモデルを使用します):
    echo 1: tiny (最小・最速)
    echo 2: base (小型・高速)
    echo 3: small (中型)
    echo 4: medium (大型 - 推奨)
    echo 5: large-v1 (旧large)
    echo 6: large-v2 (高精度)
    echo 7: large-v3 (最新)
    set /p whisper_model_choice=選択 (1-7): 
    echo.
    echo 字幕の品質を比較してYouTubeの字幕が良い場合はそちらを使用しますか？ (Y/N) (デフォルト: N):
    echo (※Nを選択すると、常にWhisper APIの結果を優先使用します)
    set /p compare_quality_choice=選択: 
    
    if /i "%compare_quality_choice%"=="Y" (
        set no_force_whisper_param=--no-force-whisper
    )
    
    set whisper_model=medium
    if "%whisper_model_choice%"=="1" set whisper_model=tiny
    if "%whisper_model_choice%"=="2" set whisper_model=base
    if "%whisper_model_choice%"=="3" set whisper_model=small
    if "%whisper_model_choice%"=="4" set whisper_model=medium
    if "%whisper_model_choice%"=="5" set whisper_model=large-v1
    if "%whisper_model_choice%"=="6" set whisper_model=large-v2
    if "%whisper_model_choice%"=="7" set whisper_model=large-v3
    
    set whisper_param=--use-whisper --whisper-api-key "%whisper_api_key%" --whisper-model %whisper_model% %no_force_whisper_param%
)

if "%image_quality%"=="" set image_quality=95
if "%text_threshold%"=="" set text_threshold=200

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
python youtube_tutorial_slides.py "%youtube_url%" --interval %interval% --lang %lang% %compact_param% --text-threshold %text_threshold% %text_split_param% --change-threshold %change_threshold% %change_detect_param% --image-quality %image_quality% --video-quality %video_quality% %whisper_param%
echo.
echo 処理が完了しました！
echo 任意のキーを押して終了...
pause > nul