import os
import cv2
import argparse
import time
import yt_dlp
from youtube_transcript_api import YouTubeTranscriptApi
from PIL import Image
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import re
from datetime import datetime

class EnhancedYouTubeTutorialExtractor:
    def __init__(self, url, output_dir="output", interval=30, lang="ja", format_type="pptx", 
                 compact_text=True, max_text_length=200, screen_change_threshold=0.6, 
                 image_quality=95, video_format="bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]"):
        """
        YouTubeチュートリアル動画から画像とテキストを抽出し、スライドファイルを生成する拡張クラス
        
        Args:
            url (str): YouTube動画のURL
            output_dir (str): 出力ディレクトリ
            interval (int): フレーム抽出間隔（秒）
            lang (str): 字幕の言語コード
            format_type (str): 出力形式 ("pptx" または "slides")
            compact_text (bool): テキストを改行せずにコンパクトに表示するか
            max_text_length (int): スライド分割の文字数閾値
            screen_change_threshold (float): 画面変化の検出閾値 (0-1の間, 高いほど敏感)
            image_quality (int): 画像の品質 (1-100, 高いほど高品質)
            video_format (str): ダウンロードする動画の形式とクオリティ
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
        self.video_id = self._extract_video_id()
        self.temp_dir = os.path.join(output_dir, "temp")
        
        # 必要なディレクトリを作成
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.temp_dir, exist_ok=True)
        
    def _extract_video_id(self):
        """YouTubeのURLからビデオIDを抽出"""
        if "youtu.be" in self.url:
            video_id = self.url.split("/")[-1]
            if "?" in video_id:
                video_id = video_id.split("?")[0]
            return video_id
        elif "youtube.com" in self.url:
            match = re.search(r"v=([^&]+)", self.url)
            if match:
                return match.group(1)
            else:
                raise ValueError("YouTubeのビデオIDが見つかりませんでした")
        else:
            raise ValueError("無効なYouTube URLです")
    
    def get_video_info(self):
        """動画のタイトルと説明を取得（yt-dlpライブラリを使用）"""
        try:
            # yt-dlpの設定オプション
            ydl_opts = {
                'skip_download': True,
                'quiet': True,
                'no_warnings': True,
            }
            
            # yt-dlpを使って情報を取得
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(self.url, download=False)
            
            # 情報を取得
            title = info.get('title', '不明なタイトル')
            author = info.get('uploader', '不明な作成者')
            description = info.get('description', '')
            length = info.get('duration', 0)
            
            # 公開日の処理
            upload_date = info.get('upload_date')
            if upload_date and len(upload_date) == 8:
                try:
                    year = int(upload_date[:4])
                    month = int(upload_date[4:6])
                    day = int(upload_date[6:8])
                    publish_date = datetime(year, month, day)
                except:
                    publish_date = None
            else:
                publish_date = None
            
            return {
                "title": title,
                "author": author,
                "description": description,
                "length": length,
                "publish_date": publish_date
            }
        except Exception as e:
            print(f"動画情報の取得に失敗しました: {e}")
            # 最低限の情報を返す
            return {
                "title": f"動画 {self.video_id}",
                "author": "不明",
                "description": "",
                "length": 0,
                "publish_date": None
            }
    
    def download_transcript(self):
        """字幕をダウンロード"""
        try:
            transcript_list = YouTubeTranscriptApi.list_transcripts(self.video_id)
            
            # 指定された言語の字幕を探す
            transcript = None
            for t in transcript_list:
                if t.language_code == self.lang:
                    transcript = t
                    break
            
            # 指定言語が見つからない場合は自動生成字幕を探す
            if transcript is None:
                for t in transcript_list:
                    if t.is_generated:
                        transcript = t
                        if t.language_code != self.lang:
                            transcript = t.translate(self.lang)
                        break
            
            # それでも見つからなければデフォルトのものを使用
            if transcript is None:
                transcript = transcript_list.find_transcript([self.lang, 'en'])
            
            # 字幕データを取得
            transcript_data = transcript.fetch()
            
            # データ形式を確認して標準化（辞書形式に統一）
            standardized_data = []
            for entry in transcript_data:
                # 既に辞書形式かチェック
                if isinstance(entry, dict) and 'start' in entry and 'text' in entry:
                    standardized_data.append({
                        'start': float(entry['start']),
                        'text': str(entry['text']),
                        'duration': float(entry.get('duration', 0))
                    })
                # オブジェクト形式の場合
                elif hasattr(entry, 'start') and hasattr(entry, 'text'):
                    standardized_data.append({
                        'start': float(entry.start),
                        'text': str(entry.text),
                        'duration': float(getattr(entry, 'duration', 0))
                    })
                # その他の場合はスキップ
                else:
                    print(f"未対応の字幕データ形式です: {type(entry)}")
            
            return standardized_data
        except Exception as e:
            print(f"字幕のダウンロードに失敗しました: {e}")
            return []
    
    def compute_frame_similarity(self, frame1, frame2):
        """2つのフレーム間の類似度を計算する (0-1, 1が完全一致)"""
        try:
            # リサイズして計算を高速化
            small1 = cv2.resize(frame1, (64, 36))
            small2 = cv2.resize(frame2, (64, 36))
            
            # グレースケールに変換
            gray1 = cv2.cvtColor(small1, cv2.COLOR_BGR2GRAY)
            gray2 = cv2.cvtColor(small2, cv2.COLOR_BGR2GRAY)
            
            # ヒストグラム比較
            hist1 = cv2.calcHist([gray1], [0], None, [256], [0, 256])
            hist2 = cv2.calcHist([gray2], [0], None, [256], [0, 256])
            
            cv2.normalize(hist1, hist1, 0, 1, cv2.NORM_MINMAX)
            cv2.normalize(hist2, hist2, 0, 1, cv2.NORM_MINMAX)
            
            # ヒストグラム類似度（コサイン距離）
            similarity = cv2.compareHist(hist1, hist2, cv2.HISTCMP_CORREL)
            
            # 構造的類似度（SSIM）も考慮したい場合はこちらを使用
            # from skimage.metrics import structural_similarity as ssim
            # ssim_score = ssim(gray1, gray2)
            # similarity = (similarity + ssim_score) / 2
            
            return max(0, min(1, similarity))  # 0-1の範囲に収める
        except Exception as e:
            print(f"類似度計算エラー: {e}")
            return 1.0  # エラーの場合は変化なしとする
    
    def extract_frames_advanced(self):
        """動画からフレームを抽出し、インターバル、テキスト量、画面変化に基づいてフレームを選別"""
        try:
            print("動画をダウンロード中...")
            video_path = os.path.join(self.temp_dir, f"{self.video_id}.mp4")
            
            # yt-dlpの設定オプション（最高品質を指定）
            ydl_opts = {
                'format': self.video_format,
                'outtmpl': video_path,
                'quiet': True,
                'no_warnings': True,
            }
            
            # yt-dlpを使って動画をダウンロード
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([self.url])
            
            if not os.path.exists(video_path):
                raise Exception("動画のダウンロードに失敗しました")
            
            print("字幕をダウンロード中...")
            transcript = self.download_transcript()
            if not transcript:
                print("警告: 字幕を取得できませんでした。テキスト分析に基づくスライド分割は無効になります。")
                
            print("フレームを抽出中...")
            cap = cv2.VideoCapture(video_path)
            fps = cap.get(cv2.CAP_PROP_FPS)
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            duration = total_frames / fps
            
            frames = []
            frame_times = []
            transcript_chunks = []
            
            # 最初のフレームを取得
            prev_frame = None
            
            # インターバルごとにフレームを抽出
            for sec in range(0, int(duration), self.interval):
                # 現在の時間ポイントを設定
                cap.set(cv2.CAP_PROP_POS_MSEC, sec * 1000)
                ret, frame = cap.read()
                
                if not ret:
                    continue
                
                # 画面変化検出（最初のフレーム以外）
                if prev_frame is not None:
                    similarity = self.compute_frame_similarity(prev_frame, frame)
                    # print(f"時間: {sec}秒, 類似度: {similarity:.2f}")
                    
                    # 類似度が閾値より低い（= 変化が大きい）場合はフレームを追加
                    if similarity < self.screen_change_threshold:
                        # 画像ファイルとして高品質で保存
                        timestamp = time.strftime('%H:%M:%S', time.gmtime(sec))
                        img_path = os.path.join(self.temp_dir, f"frame_{sec}.jpg")
                        
                        # 高画質パラメータを設定
                        img_params = [cv2.IMWRITE_JPEG_QUALITY, self.image_quality]
                        cv2.imwrite(img_path, frame, img_params)
                        
                        # フレーム情報を追加
                        frames.append(img_path)
                        frame_times.append(sec)
                        
                        # テキスト取得
                        text = self._get_transcript_at_time(transcript, sec, window_size=15)
                        transcript_chunks.append(text)
                        
                        # デバッグ出力
                        print(f"画面変化検出: {sec}秒 (類似度: {similarity:.2f})")
                
                # 定期的なインターバルでのフレーム保存（高品質）
                timestamp = time.strftime('%H:%M:%S', time.gmtime(sec))
                img_path = os.path.join(self.temp_dir, f"frame_{sec}.jpg")
                
                # 高画質パラメータを設定
                img_params = [cv2.IMWRITE_JPEG_QUALITY, self.image_quality]
                cv2.imwrite(img_path, frame, img_params)
                
                frames.append(img_path)
                frame_times.append(sec)
                
                # テキスト取得とテキスト量に基づく分割処理
                text = self._get_transcript_at_time(transcript, sec, window_size=15)
                transcript_chunks.append(text)
                
                # 定期的にテキスト量をチェックし、閾値を超えた時点で追加のフレームを取得
                current_time = sec
                if transcript and text and len(text) > self.max_text_length:
                    # テキスト量が多い場合、セグメントを分割して追加のフレームを作成
                    segments = self._split_text_by_amount(transcript, current_time, current_time + self.interval)
                    
                    for seg_time in segments[1:]:  # 最初のセグメントは既に追加済み
                        # フレームを取得
                        cap.set(cv2.CAP_PROP_POS_MSEC, seg_time * 1000)
                        ret, seg_frame = cap.read()
                        if ret:
                            # 画像ファイルとして高品質で保存
                            timestamp = time.strftime('%H:%M:%S', time.gmtime(seg_time))
                            img_path = os.path.join(self.temp_dir, f"frame_{seg_time}.jpg")
                            
                            # 高画質パラメータを設定
                            img_params = [cv2.IMWRITE_JPEG_QUALITY, self.image_quality]
                            cv2.imwrite(img_path, seg_frame, img_params)
                            
                            # フレーム情報を追加
                            frames.append(img_path)
                            frame_times.append(seg_time)
                            
                            # テキスト取得
                            seg_text = self._get_transcript_at_time(transcript, seg_time, window_size=15)
                            transcript_chunks.append(seg_text)
                            
                            # デバッグ出力
                            print(f"テキスト量による分割: {seg_time}秒")
                
                # 現在のフレームを保存
                prev_frame = frame
            
            cap.release()
            
            # フレーム時間に基づいてソート
            sorted_data = sorted(zip(frames, frame_times, transcript_chunks), key=lambda x: x[1])
            frames = [item[0] for item in sorted_data]
            frame_times = [item[1] for item in sorted_data]
            transcript_chunks = [item[2] for item in sorted_data]
            
            print(f"合計 {len(frames)} フレームを抽出しました")
            return frames, frame_times, transcript_chunks
        
        except Exception as e:
            print(f"フレーム抽出に失敗しました: {e}")
            import traceback
            traceback.print_exc()  # スタックトレースを表示
            return [], [], []
    
    def _split_text_by_amount(self, transcript, start_time, end_time):
        """指定時間範囲内の字幕を文字量に基づいて分割し、分割点の時間を返す"""
        if not transcript:
            return [start_time]
            
        # 時間範囲内の字幕エントリを抽出
        time_range_entries = [
            entry for entry in transcript 
            if start_time <= entry['start'] < end_time
        ]
        
        if not time_range_entries:
            return [start_time]
            
        # 時間順にソート
        time_range_entries.sort(key=lambda x: x['start'])
        
        # 分割点を格納するリスト（最初の時間点を含む）
        split_points = [start_time]
        
        # 文字数を累積していく
        accumulated_text = ""
        
        for entry in time_range_entries:
            accumulated_text += " " + entry['text']
            
            # 文字数が閾値を超えたら分割点として時間を追加
            if len(accumulated_text) >= self.max_text_length:
                split_points.append(entry['start'])
                accumulated_text = ""
        
        return split_points
    
    def _get_transcript_at_time(self, transcript, time_sec, window_size=30):
        """
        指定された時間付近の字幕テキストを取得
        Args:
            transcript: 字幕データ
            time_sec: 検索する時間（秒）
            window_size: 検索する時間範囲（秒）
        """
        if not transcript:
            return ""
            
        relevant_parts = []
        last_end_time = -1
        
        # 時間順に字幕をソート
        sorted_transcript = sorted(transcript, key=lambda x: x['start'])
        
        for entry in sorted_transcript:
            start_time = entry['start']
            
            # 指定時間範囲内のエントリーを探す
            if abs(start_time - time_sec) <= window_size:
                text = entry['text'].strip()
                if not text:
                    continue
                
                # コンパクトモードの場合、近い時間の字幕をまとめる
                if self.compact_text and relevant_parts and start_time - last_end_time < 2.0:
                    # 直前の字幕と結合（改行なし）
                    relevant_parts[-1] += " " + text
                else:
                    relevant_parts.append(text)
                
                last_end_time = start_time + entry.get('duration', 0)
        
        # スペースで結合（改行なし）
        return " ".join(relevant_parts) if self.compact_text else "\n".join(relevant_parts)
    
    def create_google_slides(self, frames_data, video_info):
        """GoogleSlides用プレゼンテーション（PPTX形式）を作成"""
        try:
            frames, frame_times, transcript_chunks = frames_data
            
            prs = Presentation()
            
            # タイトルスライド（GoogleSlides互換レイアウト）
            title_slide = prs.slides.add_slide(prs.slide_layouts[0])  # タイトルスライド
            
            # タイトルを設定（プレースホルダがあれば使用）
            try:
                title_placeholder = title_slide.shapes.title
                title_placeholder.text = video_info["title"]
                for paragraph in title_placeholder.text_frame.paragraphs:
                    paragraph.font.size = Pt(32)
                    paragraph.font.bold = True
                    paragraph.alignment = 1  # 中央揃え
            except:
                # プレースホルダがなければテキストボックスで追加
                left = Inches(0.5)
                top = Inches(1.0)
                width = prs.slide_width - Inches(1.0)
                height = Inches(1.0)
                
                title_box = title_slide.shapes.add_textbox(left, top, width, height)
                tf = title_box.text_frame
                p = tf.add_paragraph()
                p.text = video_info["title"]
                p.font.size = Pt(32)
                p.font.bold = True
                p.alignment = 1  # 中央揃え
            
            # サブタイトルを設定（プレースホルダがあれば使用）
            try:
                subtitle_placeholder = title_slide.placeholders[1]
                subtitle_placeholder.text = f"作成者: {video_info['author']}\n抽出日: {datetime.now().strftime('%Y-%m-%d')}"
                for paragraph in subtitle_placeholder.text_frame.paragraphs:
                    paragraph.font.size = Pt(18)
            except:
                # プレースホルダがなければテキストボックスで追加
                left = Inches(0.5)
                top = Inches(2.5)
                width = prs.slide_width - Inches(1.0)
                height = Inches(1.0)
                
                subtitle_box = title_slide.shapes.add_textbox(left, top, width, height)
                tf = subtitle_box.text_frame
                p = tf.add_paragraph()
                p.text = f"作成者: {video_info['author']}"
                p.font.size = Pt(18)
                
                p = tf.add_paragraph()
                p.text = f"抽出日: {datetime.now().strftime('%Y-%m-%d')}"
                p.font.size = Pt(18)
            
            # 目次スライド
            toc_slide = prs.slides.add_slide(prs.slide_layouts[1])  # タイトルと内容レイアウト
            
            # 目次タイトルの設定
            try:
                title_placeholder = toc_slide.shapes.title
                title_placeholder.text = "目次"
                for paragraph in title_placeholder.text_frame.paragraphs:
                    paragraph.font.size = Pt(28)
                    paragraph.font.bold = True
            except Exception as e:
                print(f"目次タイトル設定エラー: {e}")
                # プレースホルダがなければテキストボックスで追加
                title_box = toc_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.5))
                tf = title_box.text_frame
                tf.text = ""  # 空のテキスト設定で改行を防止
                p = tf.paragraphs[0]
                p.text = "目次"
                p.font.size = Pt(28)
                p.font.bold = True
            
            # 目次テキストを追加
            left = Inches(0.5)
            top = Inches(1.5)
            width = prs.slide_width - Inches(1.0)
            height = Inches(5.0)
            
            txBox = toc_slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.word_wrap = True
            
            # 既存の段落を使用して改行の追加を防止
            tf.text = ""
            p = tf.paragraphs[0]
            
            # 最初の目次項目を追加
            if frame_times and len(frame_times) > 0:
                timestamp = time.strftime('%H:%M:%S', time.gmtime(frame_times[0]))
                preview = transcript_chunks[0][:30] + "..." if transcript_chunks[0] and len(transcript_chunks[0]) > 30 else transcript_chunks[0] or "..."
                p.text = f"1. [{timestamp}] {preview}"
                
                # 残りの目次項目を追加
                for i in range(1, len(frame_times)):
                    timestamp = time.strftime('%H:%M:%S', time.gmtime(frame_times[i]))
                    preview = transcript_chunks[i][:30] + "..." if transcript_chunks[i] and len(transcript_chunks[i]) > 30 else transcript_chunks[i] or "..."
                    
                    p = tf.add_paragraph()
                    p.text = f"{i+1}. [{timestamp}] {preview}"
            else:
                p.text = "目次項目はありません"
            
            # コンテンツスライド - 単純化したレイアウトを使用
            for i, (frame_path, frame_time, transcript_text) in enumerate(zip(frames, frame_times, transcript_chunks)):
                # 空白のスライドを追加
                slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白レイアウト
                timestamp = time.strftime('%H:%M:%S', time.gmtime(frame_time))
                
                # タイトルをテキストボックスとして手動で追加（上部に配置）
                left = Inches(0.5)
                top = Inches(0.1)  # 位置を上に移動
                width = prs.slide_width - Inches(1.0)
                height = Inches(0.5)
                
                title_box = slide.shapes.add_textbox(left, top, width, height)
                tf = title_box.text_frame
                # テキストフレームのプロパティを直接設定し、改行を防止
                tf.text = ""
                p = tf.paragraphs[0]  # 最初の段落を使用
                p.text = f"セクション {i+1} - {timestamp}"
                p.font.size = Pt(20)
                p.font.bold = True
                
                # フレーム画像（スライドの80%のサイズ）
                if os.path.exists(frame_path):
                    try:
                        # スライドのサイズを取得
                        slide_width = prs.slide_width
                        slide_height = prs.slide_height
                        
                        # スライドの画質向上のため、解像度を考慮した適切なサイズ比率を計算
                        try:
                            # 画像ファイルを読み込む
                            image = cv2.imread(frame_path)
                            if image is not None:
                                frame_height, frame_width = image.shape[:2]
                                slide_aspect_ratio = 16/9  # スライドのアスペクト比（標準的なワイドスクリーン）
                                
                                # アスペクト比を考慮して配置する画像の幅を決定
                                if frame_width/frame_height > slide_aspect_ratio:
                                    # 横長の画像の場合、幅を基準に調整（横幅の80%に拡大）
                                    img_width = int(slide_width * 0.8)
                                    img_height = int(img_width * frame_height / frame_width)
                                else:
                                    # 縦長または正方形の画像の場合、高さを基準に調整（高さの60%に拡大）
                                    img_height = int(slide_height * 0.6)
                                    img_width = int(img_height * frame_width / frame_height)
                                    # 幅が横幅の80%を超える場合は調整
                                    if img_width > int(slide_width * 0.8):
                                        img_width = int(slide_width * 0.8)
                                        img_height = int(img_width * frame_height / frame_width)
                            else:
                                # 画像が読み込めない場合はデフォルト値を使用
                                img_width = int(slide_width * 0.8)
                                img_height = int(slide_height * 0.5)
                        except Exception as e:
                            print(f"画像サイズ計算エラー: {e}")
                            # エラー時はデフォルト値を使用
                            img_width = int(slide_width * 0.8)  # 横幅の80%に拡大
                            img_height = int(slide_height * 0.6)
                        
                        # 画像を適切な位置に配置
                        img_left = int((slide_width - img_width) / 2)
                        img_top = Inches(0.7)  # タイトルとの間隔を適切に設定
                        
                        pic = slide.shapes.add_picture(frame_path, img_left, img_top, width=img_width)
                        
                        # 画像の高さを取得
                        img_height = pic.height
                    except Exception as e:
                        print(f"画像の追加に失敗しました: {e}")
                        img_height = Inches(3.0)  # デフォルト値
                else:
                    img_height = Inches(3.0)  # 画像がない場合のデフォルト値
                
                # トランスクリプトテキスト（画像の後ろにかからないように下部に配置）
                if transcript_text:
                    left = Inches(0.5)
                    text_top = Inches(0.8) + img_height + Inches(0.2)  # 画像の下に余白を設けて配置
                    width = prs.slide_width - Inches(1.0)
                    height = Inches(1.5)
                    
                    text_box = slide.shapes.add_textbox(left, text_top, width, height)
                    tf = text_box.text_frame
                    tf.word_wrap = True
                    p = tf.add_paragraph()
                    p.text = transcript_text
                    p.font.size = Pt(14)
                
                # YouTube リンク（スライドの最下部に配置、クリック可能なハイパーリンクとして）
                youtube_link = f"{self.url}&t={int(frame_time)}s"
                link_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(9), Inches(0.4))
                tf = link_box.text_frame
                p = tf.add_paragraph()
                run = p.add_run()
                run.text = "この部分を動画で見る"
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(0, 0, 255)
                run.hyperlink.address = youtube_link
            
            # PPTXを保存（GoogleSlides用）
            safe_title = ''.join(c for c in video_info["title"] if c.isalnum() or c.isspace() or c == '_')[:30]
            pptx_path = os.path.join(self.output_dir, f"{safe_title}_googleslides.pptx")
            prs.save(pptx_path)
            print(f"GoogleSlides用ファイルを作成しました: {pptx_path}")
            return pptx_path
        
        except Exception as e:
            print(f"GoogleSlides用ファイル作成に失敗しました: {e}")
            import traceback
            traceback.print_exc()  # 詳細なエラー情報を表示
            return None
    
    def process(self):
        """メイン処理"""
        print(f"YouTube URL: {self.url}の処理を開始します")
        
        # 動画情報を取得
        video_info = self.get_video_info()
        print(f"動画タイトル: {video_info['title']}")
        
        # 拡張機能付きフレーム抽出（テキスト量・画面変化による分割を含む）
        frames_data = self.extract_frames_advanced()
        if not frames_data[0]:
            print("フレームを抽出できませんでした。処理を中止します。")
            return None
        
        result_path = None
        
        # GoogleSlides形式でスライドを生成
        result_path = self.create_google_slides(frames_data, video_info)
        
        # 一時ファイルを削除
        for frame_path in frames_data[0]:
            if os.path.exists(frame_path):
                os.remove(frame_path)
        
        if os.path.exists(os.path.join(self.temp_dir, f"{self.video_id}.mp4")):
            os.remove(os.path.join(self.temp_dir, f"{self.video_id}.mp4"))
        
        return result_path
