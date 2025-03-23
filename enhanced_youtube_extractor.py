"""YouTube動画からキャプチャとテキストを抽出するモジュール。

このモジュールはYouTube動画から定期的に画像をキャプチャし、
対応する字幕とともにスライドに変換します。
"""

import os
import re
import time
from datetime import datetime

import cv2
import numpy as np
import yt_dlp
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from youtube_transcript_api import YouTubeTranscriptApi

# 追加: Whisper API連携モジュールをインポート
try:
    from whisper_integration import WhisperTranscriptionProvider, TranscriptQualitySelector
    WHISPER_AVAILABLE = True
except ImportError:
    WHISPER_AVAILABLE = False


class EnhancedYouTubeTutorialExtractor:
    """YouTubeチュートリアル動画から画像とテキストを抽出し、スライドを生成するクラス。"""

    def __init__(
            self,
            url,
            output_dir="output",
            interval=30,
            lang="ja",
            format_type="pptx",
            compact_text=True,
            max_text_length=200,
            screen_change_threshold=0.6,
            image_quality=95,
            video_format="bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]",
            add_thumbnail=True,
            use_whisper=False,
            whisper_api_key=None,
            whisper_model="medium",
            force_whisper=False):
        """初期化メソッド。

        Args:
            url: YouTube動画のURL
            output_dir: 出力ディレクトリ
            interval: フレーム抽出間隔（秒）
            lang: 字幕の言語コード
            format_type: 出力形式 ("pptx" または "slides")
            compact_text: テキストを改行せずにコンパクトに表示するか
            max_text_length: スライド分割の文字数閾値
            screen_change_threshold: 画面変化の検出閾値 (0-1の間, 高いほど敏感)
            image_quality: 画像の品質 (1-100, 高いほど高品質)
            video_format: ダウンロードする動画の形式とクオリティ
            add_thumbnail: サムネイルを追加するかどうか
            use_whisper: Whisper APIを使用するかどうか
            whisper_api_key: Whisper API（OpenAI API）のキー
            whisper_model: 使用するWhisperモデル
            force_whisper: 品質評価にかかわらず常にWhisper APIの結果を使用するか
        """
        self.url = url
        self.output_dir = output_dir
        self.interval = interval
        self.lang = lang
        self.format_type = format_type
        self.compact_text = compact_text
        self.max_text_length = max_text_length
        self.screen_change_threshold = screen_change_threshold
        self.image_quality = image_quality
        self.video_format = video_format
        self.add_thumbnail = add_thumbnail
        self.video_id = self._extract_video_id()
        self.temp_dir = os.path.join(output_dir, "temp")
        
        # 追加: Whisper API関連の設定
        self.use_whisper = use_whisper and WHISPER_AVAILABLE
        self.whisper_api_key = whisper_api_key
        self.whisper_model = whisper_model
        self.force_whisper = force_whisper
        
        # 必要なディレクトリを作成
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.temp_dir, exist_ok=True)