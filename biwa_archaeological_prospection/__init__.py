"""
琵琶湖考古学的探査パイプライン
Lake Biwa Archaeological Prospection Pipeline

地形データ（DEM）から遺跡候補地を検出するための解析パイプライン。
国土地理院の公開タイルデータを使用し、以下の手法を適用:
- 赤色立体地図 (RRIM: Red Relief Image Map)
- 天空率 (SVF: Sky-View Factor)
- 陰影起伏図 (Hillshade)
- 傾斜量図 (Slope)
- 局所起伏モデル (LRM: Local Relief Model)
- 統計的異常検出
- 古琵琶湖変遷の地質学的コンテキスト
"""

__version__ = "0.1.0"
