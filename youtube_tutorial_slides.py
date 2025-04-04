#!/usr/bin/env python3
"""YouTubeチュートリアル動画からスライドを生成するツール。

このスクリプトはYouTube動画のURLからキャプチャとテキストを抽出し、
PowerPointのスライドに変換します。
"""

import argparse
import os

from enhanced_youtube_extractor import EnhancedYouTubeTutorialExtractor


def main():
    """コマンドライン引数を解析し、スライド生成処理を実行する。"""
    parser = argparse.ArgumentParser(description="YouTube チュートリアル動画からスライドを生成（拡張版）")
    parser.add_argument("url", help="YouTube 動画の URL")
    parser.add_argument("--output", "-o", default="output", help="出力ディレクトリ")
    parser.add_argument(
        "--format", "-f", choices=["pptx", "slides"], default="slides", help="出力形式")
    parser.add_argument("--interval", "-i", type=int, default=30, help="フレーム抽出間隔（秒）")
    parser.add_argument("--lang", "-l", default="ja", help="字幕の言語コード（例: ja, en）")
    parser.add_argument(
        "--nocompact", "-nc", action="store_false", help="テキストを改行せずにまとめない")
    
    # 拡張機能のための新しいパラメータ
    parser.add_argument(
        "--text-threshold", "-tt", type=int, default=200, 
        help="スライド分割を行う文字数閾値 (デフォルト: 200文字)")
    parser.add_argument(
        "--change-threshold", "-ct", type=float, default=0.6, 
        help="画面変化検出の閾値 (0-1, 低いほど敏感。デフォルト: 0.6)")
    parser.add_argument(
        "--disable-text-split", "-dts", action="store_true",
        help="文字量によるスライド分割を無効にする")
    parser.add_argument(
        "--disable-change-detection", "-dcd", action="store_true",
        help="画面変化検出によるスライド分割を無効にする")
    parser.add_argument(
        "--image-quality", "-iq", type=int, default=95, 
        help="画像品質 (1-100, 高いほど高品質。デフォルト: 95)")
    parser.add_argument(
        "--video-quality", "-vq", choices=["best", "high", "medium", "low"], default="high", 
        help="動画ダウンロード品質 (best=最高品質, high=高品質, medium=中品質, low=低品質。デフォルト: high)")
    parser.add_argument(
        "--no-thumbnail", action="store_true",
        help="最初のスライドにYouTubeサムネイルを追加しない")
    
    # Whisper API 関連のパラメータ
    parser.add_argument(
        "--use-whisper", action="store_true",
        help="Whisper APIを使用して高精度な音声認識を行う (デフォルトで字幕優先モードになります)")
    parser.add_argument(
        "--whisper-api-key", 
        help="Whisper API（OpenAI API）のキー。環境変数 OPENAI_API_KEY からも取得可能")
    parser.add_argument(
        "--whisper-model", 
        choices=["tiny", "base", "small", "medium", "large-v1", "large-v2", "large-v3"], 
        default="medium",
        help="使用するWhisperモデル（注: 現在のOpenAI APIではすべて内部的に同じモデルを使用します）")
    parser.add_argument(
        "--no-force-whisper", action="store_true",
        help="Whisper API使用時に品質比較を行い、必ずしもWhisper APIの結果を優先しないようにする")
    
    args = parser.parse_args()
    
    # 文字量閾値の調整（無効化オプションが指定されていれば最大値に設定）
    text_threshold = 1000000 if args.disable_text_split else args.text_threshold
    
    # 画面変化閾値の調整（無効化オプションが指定されていれば0に設定）
    change_threshold = 0.0 if args.disable_change_detection else args.change_threshold
    
    # 動画品質設定の変換
    video_format_map = {
        "best": "bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]",
        "high": "bestvideo[height<=1080][ext=mp4]+bestaudio[ext=m4a]/best[height<=1080][ext=mp4]",
        "medium": "bestvideo[height<=720][ext=mp4]+bestaudio[ext=m4a]/best[height<=720][ext=mp4]",
        "low": "bestvideo[height<=480][ext=mp4]+bestaudio[ext=m4a]/best[height<=480][ext=mp4]"
    }
    video_format = video_format_map.get(args.video_quality, video_format_map["high"])
    
    # Whisper APIキーの取得（コマンドライン引数 > 環境変数）
    whisper_api_key = args.whisper_api_key
    if args.use_whisper and not whisper_api_key:
        whisper_api_key = os.environ.get('OPENAI_API_KEY')
        if not whisper_api_key:
            print("警告: Whisper APIが有効ですが、APIキーが設定されていません。")
            print("--whisper-api-key オプションまたは環境変数 OPENAI_API_KEY を設定してください。")
            print("Whisper APIを使用せずに処理を続行します。")
            args.use_whisper = False
    
    # デフォルトで force_whisper を有効に（--no-force-whisper が指定されていない場合）
    force_whisper = args.use_whisper and not args.no_force_whisper
    
    extractor = EnhancedYouTubeTutorialExtractor(
        url=args.url,
        output_dir=args.output,
        format_type=args.format,
        interval=args.interval,
        lang=args.lang,
        compact_text=args.nocompact,
        max_text_length=text_threshold,
        screen_change_threshold=change_threshold,
        image_quality=args.image_quality,
        video_format=video_format,
        add_thumbnail=not args.no_thumbnail,
        use_whisper=args.use_whisper,
        whisper_api_key=whisper_api_key,
        whisper_model=args.whisper_model,
        force_whisper=force_whisper
    )
    
    result_path = extractor.process()
    
    if result_path:
        print("\n処理が完了しました！")
        print(f"生成ファイル: {result_path}")
    else:
        print("処理に失敗しました。")


if __name__ == "__main__":
    main()
