"""Whisper APIと連携して音声認識の精度を高めるモジュール。

このモジュールはYouTube動画から音声を抽出し、Whisper APIを使用して
より高精度な字幕を生成します。
"""

import os
import tempfile
import subprocess
import time
import requests
from typing import List, Dict, Any, Optional


class WhisperTranscriptionProvider:
    """Whisper APIを利用して音声から字幕を生成するクラス。"""
    
    def __init__(self, api_key: str, language: str = "ja", model: str = "medium"):
        """初期化メソッド。
        
        Args:
            api_key: OpenAI APIキー
            language: 字幕の言語コード (例: "ja", "en")
            model: 使用するWhisperモデル名（UIの互換性のために保持、内部的にはwhisper-1を使用）
        """
        self.api_key = api_key
        self.language = language
        self.model = model  # 現在はUIの互換性のために保持
        self.api_url = "https://api.openai.com/v1/audio/transcriptions"
        
    def extract_audio(self, video_path: str, output_dir: str) -> Optional[str]:
        """動画から音声を抽出する。
        
        Args:
            video_path: 動画ファイルのパス
            output_dir: 出力ディレクトリ
            
        Returns:
            str: 音声ファイルのパス。失敗時はNone。
        """
        # 出力ファイル名を設定
        audio_filename = os.path.splitext(os.path.basename(video_path))[0] + ".mp3"
        audio_path = os.path.join(output_dir, audio_filename)
        
        # FFmpegを使用して音声を抽出
        try:
            command = [
                "ffmpeg", "-i", video_path, 
                "-q:a", "0", "-map", "a", "-vn", 
                audio_path
            ]
            subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            return audio_path
        except subprocess.CalledProcessError as e:
            print(f"音声抽出エラー: {e}")
            return None
        except Exception as e:
            print(f"予期しないエラー: {e}")
            return None
            
    def _split_audio(self, audio_path: str, output_dir: str, segment_length: int = 600) -> List[str]:
        """音声ファイルを指定された長さのセグメントに分割する。
        
        Args:
            audio_path: 音声ファイルのパス
            output_dir: 出力ディレクトリ
            segment_length: セグメントの長さ（秒）
            
        Returns:
            List[str]: 分割された音声ファイルのパスリスト
        """
        # 音声の長さを取得
        try:
            command = ["ffprobe", "-v", "error", "-show_entries", "format=duration", 
                      "-of", "default=noprint_wrappers=1:nokey=1", audio_path]
            result = subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            duration = float(result.stdout)
        except Exception as e:
            print(f"音声の長さを取得できませんでした: {e}")
            return [audio_path]  # 分割せずに元のファイルを返す
            
        # セグメント数を計算
        num_segments = int(duration / segment_length) + 1
        segment_files = []
        
        # 音声を分割
        base_name = os.path.splitext(os.path.basename(audio_path))[0]
        
        for i in range(num_segments):
            start_time = i * segment_length
            segment_file = os.path.join(output_dir, f"{base_name}_segment_{i:03d}.mp3")
            
            try:
                command = [
                    "ffmpeg", "-i", audio_path, 
                    "-ss", str(start_time), 
                    "-t", str(segment_length),
                    "-c:a", "libmp3lame", "-q:a", "4", 
                    segment_file
                ]
                subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                segment_files.append(segment_file)
            except Exception as e:
                print(f"セグメント {i} の分割に失敗しました: {e}")
                
        return segment_files
            
    def transcribe_audio(self, audio_path: str) -> List[Dict[str, Any]]:
        """Whisper APIを使用して音声を文字起こしする。
        
        Args:
            audio_path: 音声ファイルのパス
            
        Returns:
            List[Dict[str, Any]]: 文字起こし結果（YouTube-Transcript-APIと互換性のある形式）
        """
        # 音声ファイルが大きい場合は分割（WhisperAPIの制限は25MB）
        # ファイルサイズを確認
        file_size = os.path.getsize(audio_path) / (1024 * 1024)  # MB単位
        
        if file_size > 24:
            print(f"音声ファイルが大きいため分割します（{file_size:.2f}MB）")
            temp_dir = tempfile.mkdtemp()
            audio_segments = self._split_audio(audio_path, temp_dir)
        else:
            audio_segments = [audio_path]
            
        all_transcripts = []
        current_start_time = 0.0
        
        # 現在のOpenAI APIではすべてwhisper-1モデルを使用
        api_model = "whisper-1"
        
        # 各セグメントを文字起こし
        for segment_path in audio_segments:
            try:
                headers = {
                    "Authorization": f"Bearer {self.api_key}"
                }
                
                with open(segment_path, "rb") as audio_file:
                    files = {
                        "file": audio_file,
                    }
                    data = {
                        "model": api_model,  # 常に "whisper-1" を使用
                        "language": self.language,
                        "response_format": "verbose_json",
                        "timestamp_granularities": ["segment"]
                    }
                    
                    # APIリクエスト
                    response = requests.post(
                        self.api_url,
                        headers=headers,
                        files=files,
                        data=data
                    )
                    
                    if response.status_code != 200:
                        print(f"API エラー: {response.status_code}")
                        print(response.text)
                        continue
                        
                    result = response.json()
                    
                    # セグメントごとの結果を変換
                    if "segments" in result:
                        for segment in result["segments"]:
                            transcript_entry = {
                                "start": current_start_time + segment.get("start", 0),
                                "text": segment.get("text", "").strip(),
                                "duration": segment.get("end", 0) - segment.get("start", 0)
                            }
                            all_transcripts.append(transcript_entry)
                    
                    # 現在の開始時間を更新
                    if audio_segments.index(segment_path) < len(audio_segments) - 1:
                        # 最後のセグメント以外は次のセグメントの開始時間を調整
                        current_start_time += 600  # 10分（600秒）ごとのセグメント
                    
            except Exception as e:
                print(f"文字起こしエラー: {e}")
                
            # レート制限対策
            time.sleep(1)
                
        # 一時ディレクトリを削除
        if file_size > 24:
            for segment_path in audio_segments:
                try:
                    os.remove(segment_path)
                except:
                    pass
            try:
                os.rmdir(temp_dir)
            except:
                pass
                
        return all_transcripts

    def get_transcript(self, video_path: str, output_dir: str) -> List[Dict[str, Any]]:
        """動画から音声を抽出し、Whisper APIで文字起こしする。
        
        Args:
            video_path: 動画ファイルのパス
            output_dir: 出力ディレクトリ
            
        Returns:
            List[Dict[str, Any]]: 文字起こし結果
        """
        # 音声抽出
        audio_path = self.extract_audio(video_path, output_dir)
        if not audio_path:
            return []
            
        # 文字起こし
        transcript = self.transcribe_audio(audio_path)
        
        # 音声ファイルを削除
        try:
            os.remove(audio_path)
        except:
            pass
            
        return transcript


class TranscriptQualitySelector:
    """複数の字幕ソースから最適な字幕を選択するクラス。"""
    
    @staticmethod
    def evaluate_transcript_quality(transcript: List[Dict[str, Any]]) -> float:
        """字幕の品質スコアを計算する。
        
        Args:
            transcript: 評価する字幕データ
            
        Returns:
            float: 品質スコア（高いほど良い）
        """
        if not transcript:
            return 0.0
            
        # 評価指標
        total_length = 0
        word_count = 0
        segment_count = len(transcript)
        
        for entry in transcript:
            text = entry.get("text", "")
            total_length += len(text)
            words = text.split()
            word_count += len(words)
            
        # 平均文字数が多すぎる/少なすぎる場合はペナルティ
        avg_length = total_length / segment_count if segment_count > 0 else 0
        length_penalty = 1.0
        if avg_length < 10:  # 短すぎる
            length_penalty = avg_length / 10
        elif avg_length > 100:  # 長すぎる
            length_penalty = 100 / avg_length
            
        # 単語の平均文字数（短すぎると不自然）
        avg_word_length = total_length / word_count if word_count > 0 else 0
        word_length_score = min(1.0, avg_word_length / 3)
        
        # 総合スコア
        quality_score = (
            0.5 * length_penalty +  # 適切な長さ
            0.3 * word_length_score +  # 単語の長さ
            0.2 * min(1.0, segment_count / 100)  # セグメント数の十分さ
        )
        
        return quality_score
        
    @staticmethod
    def select_best_transcript(youtube_transcript: List[Dict[str, Any]], 
                               whisper_transcript: List[Dict[str, Any]],
                               force_whisper: bool = False) -> List[Dict[str, Any]]:
        """2つの字幕ソースから品質の高い方を選択する。
        
        Args:
            youtube_transcript: YouTubeから取得した字幕
            whisper_transcript: Whisper APIで生成した字幕
            force_whisper: Whisperの結果を強制的に選択するフラグ
            
        Returns:
            List[Dict[str, Any]]: 選択された字幕
        """
        # Whisper APIの結果を強制的に選択する場合
        if force_whisper:
            # Whisper APIの結果がある場合はそれを使用
            if whisper_transcript:
                print("Whisper APIの字幕を使用します（強制優先設定）")
                return whisper_transcript
            # Whisper APIの結果がない場合はYouTubeの字幕を使用
            elif youtube_transcript:
                print("Whisper APIの結果が取得できなかったため、YouTubeの字幕を使用します")
                return youtube_transcript
            else:
                return []
                
        # 通常の品質比較による選択
        youtube_score = TranscriptQualitySelector.evaluate_transcript_quality(youtube_transcript)
        whisper_score = TranscriptQualitySelector.evaluate_transcript_quality(whisper_transcript)
        
        print(f"字幕評価スコア - YouTube: {youtube_score:.2f}, Whisper: {whisper_score:.2f}")
        
        if whisper_score > youtube_score:
            print("Whisper APIの字幕を使用します（品質が高いと判断）")
            return whisper_transcript
        else:
            print("YouTubeの字幕を使用します（品質が高いと判断）")
            return youtube_transcript
