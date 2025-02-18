from aip import AipSpeech
from config import Config

class BaiduVoiceAPI:
    def __init__(self):
        self.client = AipSpeech(
            Config.BAIDU_APP_ID,
            Config.BAIDU_API_KEY,
            Config.BAIDU_SECRET_KEY
        )

    def recognize(self, audio_data):
        result = self.client.asr(audio_data, 'pcm', 16000, {
            'dev_pid': 1537,  # Mandarin
        })

        if result['err_no'] == 0:
            return result['result'][0]
        return ''