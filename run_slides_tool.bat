@echo off
echo YouTube動画からGoogleスライド作成ツール
echo ========================================
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
echo 処理を開始します...
python youtube_tutorial_slides.py "%youtube_url%" --interval %interval% --lang %lang% %compact_param%
echo.
echo 処理が完了しました！
echo 任意のキーを押して終了...
pause > nul