import sys
import os
import datetime

import pyauto
from keyhac import *


def configure(keymap):
    # 指定された条件でキーマップを取得
    keymap_app = keymap.defineWindowKeymap("EXCEL.exe")

    # ホットキーを定義
    hotkeys = {"C-E": "F2",               # Ctrl+E を F2 に変更
               "C-Right": ["A-H", "R"],   # 複数キーのシーケンスに置換する場合はリストにする
               }
    # ホットキーの内容をキーマップに設定
    for key, value in hotkeys.items():
        keymap_app[key] = value
        