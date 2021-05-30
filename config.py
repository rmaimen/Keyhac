import os
import re
import datetime
import time
from keyhac import *

def configure(keymap):

    # scoop で入手した vscode で編集（なければメモ帳）
    EDITOR_PATH = r"C:\Users\{}\scoop\apps\vscode\current\Code.exe".format(os.environ.get("USERNAME"))
    if not os.path.exists(EDITOR_PATH):
        EDITOR_PATH = "notepad.exe"
    keymap.editor = EDITOR_PATH

    # theme
    keymap.setFont("HackGen Console", 16)
    keymap.setTheme("black")

    # 全体用キー設定の読み込み
    setting_global(keymap)

    # Excel用キー設定の読み込み
    setting_excel(keymap)

    # Osqledit用キー設定の読み込み
    setting_osqledit(keymap)


def setting_global(keymap):
    def global_dateStr():
        setClipboardText("")
        delay(0.1)
        clipboard_str = copy_string(keymap)

        result_str = date_string_format_change(clipboard_str)

        keymap.InputTextCommand(result_str)()

        # 入力部分を選択状態にする
        for i in result_str:
            keymap.InputKeyCommand("S-Left")()

    
    def reload():
        keymap.command_ReloadConfig()


    # 全体で有効なキーマップのオブジェクト
    keymap_global = keymap.defineWindowKeymap()

    # ホットキーを定義
    hotkeys = {                        
        # 設定ファイルのリロード
        "C-A-R": lambda: reload(),
        # 拡張キー(F13(124)～F21(132))
        "(124)": "Esc", 
        # 日付入力
        "(125)": lambda: global_dateStr(),
        "(126)":"Esc",
        "(127)":"Esc",
        "(128)":"Esc",
        "(129)":"Esc",
        "(130)":"Esc",
        "(131)":"Esc",
        "(132)":"Esc"
    }

    # ホットキーの内容をキーマップに設定
    for key, value in hotkeys.items():
        keymap_global[key] = value


def setting_excel(keymap):
    def excel_select_row():
        send_input(keymap, ["S-Space"])

    def excel_dateStr():
        keys = []
        setClipboardText("")
        keys.append("F2")
        keys.append("S-Home")
        
        send_input(keymap,keys, False, 0.1)

        clipboard_str = copy_string(keymap)
        
        result_str = date_string_format_change(clipboard_str)

        keymap.InputTextCommand(result_str)()
        delay(0.2)

        keys.clear()
        keys.append("Enter")
        keys.append("Up")
        send_input(keymap,keys, False, 0.1)

    # 指定された条件でキーマップを取得
    keymap_excel = keymap.defineWindowKeymap("EXCEL.exe")

    # ホットキーを定義
    hotkeys = {
        # F1 ヘルプを無効化
        "F1": "Esc",                         
        # IMEがONでも、行選択(Ctrl+Space)を行えるようにする
        "S-Space": lambda: excel_select_row(),
        # 拡張キー(F13(124)～F21(132))
        # 画面固定トグル
        "(124)":["A-W","F","F","Esc", "A-H", "Enter"],
        # 日付入力
        "(125)": lambda: excel_dateStr(),
        # エラー検索を割り当て
        "(126)":["C-G","A-S","F","U","X","G","Enter"], 
        # 表示範囲選択
        "(127)":"A-Semicolon",
        # 参照先トレース
        "(128)":["A-M","D"],
        # トレース線削除
        "(129)":["A-M","A","A"],
        # コピー
        "(130)":"C-C",
        # 値貼り付け
        "(131)":["C-A-V","V","Enter", "A-H", "Enter"],
        # 書式貼り付け
        "(132)":["C-A-V","T","Enter", "A-H", "Enter"],
    }

    # ホットキーの内容をキーマップに設定
    for key, value in hotkeys.items():
        keymap_excel[key] = value


def setting_osqledit(keymap):
    def edata_sql_template():
        ymd = datetime.date.today().strftime(r"%Y%m%d")

        # keys = []
        # keys.append("SELECT")
        # keys.append("Enter")
        # keys.append("\t*")
        # keys.append("Enter")
        # keys.append("FROM")
        # keys.append("Enter")
        # keys.append("\tD_EDATA E")
        # keys.append("Enter")
        # keys.append("WHERE 1=1")
        # keys.append("Enter")
        # keys.append("\tAND E.ED_SYUYMD = " + ymd)
        # send_input(keymap, keys, False, 0.1)

        sql_str = ""
        sql_str += "SELECT\n"
        sql_str += "\t*\n"
        sql_str += "FROM\n"
        sql_str += "\tD_EDATA E\n"
        sql_str += "WHERE 1=1\n"
        sql_str += "\tAND E.ED_SYUYMD >= " + ymd + "\n"
        sql_str += "\tAND E.ED_SYUYMD <= " + ymd + "\n"
        sql_str += "\tAND E.ED_EOCD = \n"
        sql_str += "\tAND E.ED_EOSSY = \n"

        paste_string(keymap, sql_str)

    def jdata_sql_template():
        ymd = datetime.date.today().strftime(r"%Y%m%d")

        sql_str = ""
        sql_str += "SELECT\n"
        sql_str += "\t*\n"
        sql_str += "FROM\n"
        sql_str += "\tD_JDATAH H\n"
        sql_str += "\tINNER JOIN D_JDATAM M\n"
        sql_str += "\t\tON H.JDH_DENYMD = M.JDM_DENYMD\n"
        sql_str += "\t\tAND H.JDH_DENNO = M.JDM_DENNO\n"
        sql_str += "\t\tAND H.JDH_SEQNO = M.JDM_SEQNO\n"
        sql_str += "WHERE 1=1\n"
        sql_str += "\tAND H.JDH_TCKYMD >= " + ymd + "\n"
        sql_str += "\tAND H.JDH_TCKYMD <= " + ymd + "\n"
        sql_str += "\tAND M.JDM_YOKAKUTEI = 2\n"

        paste_string(keymap, sql_str)

    def udata_sql_template():
        ymd = datetime.date.today().strftime(r"%Y%m%d")

        sql_str = ""
        sql_str += "SELECT\n"
        sql_str += "\t*\n"
        sql_str += "FROM\n"
        sql_str += "\tD_UDATAH H\n"
        sql_str += "\tINNER JOIN D_UDATAM M\n"
        sql_str += "\t\tON H.UDH_DENYMD = M.UDM_DENYMD\n"
        sql_str += "\t\tAND H.UDH_DENNO = M.UDM_DENNO\n"
        sql_str += "\t\tAND H.UDH_SEQNO = M.UDM_SEQNO\n"
        sql_str += "WHERE 1=1\n"
        sql_str += "\tAND H.UDH_SEIYMD >= " + ymd + "\n"
        sql_str += "\tAND H.UDH_SEIYMD <= " + ymd + "\n"

        paste_string(keymap, sql_str)

    def judata_sql_template():
        ymd = datetime.date.today().strftime(r"%Y%m%d")

        sql_str = ""
        sql_str += "SELECT\n"
        sql_str += "\t*\n"
        sql_str += "FROM\n"
        sql_str += "\t(\n"
        sql_str += "\tSELECT\n"
        sql_str += "\t\t*\n"
        sql_str += "\tFROM\n"
        sql_str += "\t\tD_JDATAH H\n"
        sql_str += "\t\tINNER JOIN D_JDATAM M\n"
        sql_str += "\t\t\tON H.JDH_DENYMD = M.JDM_DENYMD\n"
        sql_str += "\t\t\tAND H.JDH_DENNO = M.JDM_DENNO\n"
        sql_str += "\t\t\tAND H.JDH_SEQNO = M.JDM_SEQNO\n"
        sql_str += "\t)J\n"
        sql_str += "\tINNER JOIN\n"
        sql_str += "\t(\n"
        sql_str += "\tSELECT\n"
        sql_str += "\t\t*\n"
        sql_str += "\tFROM\n"
        sql_str += "\t\tD_UDATAH H\n"
        sql_str += "\t\tINNER JOIN D_UDATAM M\n"
        sql_str += "\t\t\tON H.UDH_DENYMD = M.UDM_DENYMD\n"
        sql_str += "\t\t\tAND H.UDH_DENNO = M.UDM_DENNO\n"
        sql_str += "\t\t\tAND H.UDH_SEQNO = M.UDM_SEQNO\n"
        sql_str += "\t)U\n"
        sql_str += "\t\tON U.UDH_JDDENYMD = J.JDH_DENYMD\n"
        sql_str += "\t\tAND U.UDH_JDDENNO = J.JDH_DENNO\n"
        sql_str += "\t\tAND U.UDH_JDSEQNO = J.JDH_SEQNO\n"
        sql_str += "\t\tAND U.UDM_GYONO = J.JDM_GYONO\n"
        sql_str += "WHERE 1=1\n"
        sql_str += "\tAND U.UDH_SEIYMD >= " + ymd + "\n"
        sql_str += "\tAND U.UDH_SEIYMD <= " + ymd + "\n"

        paste_string(keymap, sql_str)
        
    # 指定された条件でキーマップを取得
    keymap_excel = keymap.defineWindowKeymap("osqledit.exe")

    # ホットキーを定義
    hotkeys = {
        # 拡張キー(F13(124)～F21(132))
        # EDATA Template
        "(124)": lambda: edata_sql_template(),
        # JDATA Template
        "(125)": lambda: jdata_sql_template(),
        # Udata Template
        "(126)": lambda: udata_sql_template(),
        # JDATA+UDATA Template
        "(127)": lambda: judata_sql_template(),
        # 
        "(128)":"Esc",
        # SQL実行
        "(129)":"C-R",
        # カラム付きコピー
        "(130)":"C-S-A-C",
        # コメント
        "(131)":"C-Slash",
        # デコメント
        "(132)":"C-S-Slash",
    }

    # ホットキーの内容をキーマップに設定
    for key, value in hotkeys.items():
        keymap_excel[key] = value

def delay(sec = 0.05):
    time.sleep(sec)

def get_clippedText():
    return (getClipboardText() or "")

def paste_string(keymap, s):
    setClipboardText(s)
    delay()
    keymap.InputKeyCommand("C-V")()

def copy_string(keymap, sec = 0.05):
    send_input(keymap, ["C-C"], sleep=sec)
    return get_clippedText()

# キー入力・テキスト入力を区別せず入力
def send_input(keymap, keys, ime_mode = False, sleep = 0.01):
    def input_command():
        for key in keys:
            try:
                keymap.InputKeyCommand(key)()
            except:
                # 動いていない？？
                keymap.InputTextCommand(key)()
            finally:
                delay(sleep)

    if (ime_mode is not None) and (keymap.getWindow().getImeStatus() != ime_mode):
        keymap.InputKeyCommand("(243)")()
        input_command()
        keymap.InputKeyCommand("(243)")()
    else:
        input_command()

# 日付入力を行う、繰り返し呼ぶことで書式を変える
def date_string_format_change(date_str):
    result_str = ""
    print(date_str)
    if date_str == "":
        # 文字列がセットされていない場合
        result_str = datetime.datetime.now().strftime(r"%Y%m%d")
    else:
        # 日付文字の場合はフォーマット変換
        pattern_dic = {
            "yyyyMMdd": r"^(?!([02468][1235679]|[13579][01345789])000229)(([0-9]{4}(0?1|0?3|0?5|0?7|0?8|10|12)(0?[1-9]|[12][0-9]|3[01]))|([0-9]{4}(0?4|0?6|0?9|11)(0?[1-9]|[12][0-9]|30))|([0-9]{4}0?2(0?[1-9]|1[0-9]|2[0-8]))|([0-9]{2}([02468][048]|[13579][26])0229))$",
            "yyyy/MM/dd": r"^(?!([02468][1235679]|[13579][01345789])00\/02\/29)(([0-9]{4}\/(0?1|0?3|0?5|0?7|0?8|10|12)\/(0?[1-9]|[12][0-9]|3[01]))|([0-9]{4}\/(0?4|0?6|0?9|11)\/(0?[1-9]|[12][0-9]|30))|([0-9]{4}\/0?2\/(0?[1-9]|1[0-9]|2[0-8]))|([0-9]{2}([02468][048]|[13579][26])\/02\/29))$",
            "yyyy-MM-dd": r"^(?!([02468][1235679]|[13579][01345789])00-02-29)(([0-9]{4}-(01|03|05|07|08|10|12)-(0[1-9]|[12][0-9]|3[01]))|([0-9]{4}-(04|06|09|11)-(0[1-9]|[12][0-9]|30))|([0-9]{4}-02-(0[1-9]|1[0-9]|2[0-8]))|([0-9]{2}([02468][048]|[13579][26])-02-29))$",
            "yyyy年MM月dd日": r"^(?!([02468][1235679]|[13579][01345789])00年02月29日)(([0-9]{4}年(01|03|05|07|08|10|12)月(0[1-9]|[12][0-9]|3[01])日)|([0-9]{4}年(04|06|09|11)月(0[1-9]|[12][0-9]|30)日)|([0-9]{4}年02月(0[1-9]|1[0-9]|2[0-8])日)|([0-9]{2}([02468][048]|[13579][26])年02月29日))$",
            "yyyy年MM月dd日(ddd)": r"^(?!([02468][1235679]|[13579][01345789])00年02月29日)(([0-9]{4}年(01|03|05|07|08|10|12)月(0[1-9]|[12][0-9]|3[01])日)|([0-9]{4}年(04|06|09|11)月(0[1-9]|[12][0-9]|30)日)|([0-9]{4}年02月(0[1-9]|1[0-9]|2[0-8])日)|([0-9]{2}([02468][048]|[13579][26])年02月29日))\([日月火水木金土]\)$",
            "MMdd": r"((01|03|05|07|08|10|12)(0[1-9]|[12][0-9]|3[01])|(04|06|09|11)(0[1-9]|[12][0-9]|30)|(02)(0[1-9]|[12][0-9]))$",
            "MM/dd,M/d": r"((0?1|0?3|0?5|0?7|0?8|10|12)\/(0?[1-9]|[12][0-9]|3[01])|(0?4|0?6|0?9|11)\/(0?[1-9]|[12][0-9]|30)|(0?2)\/(0?[1-9]|[12][0-9]))$",
            "MM-dd,M-d": r"((0?1|0?3|0?5|0?7|0?8|10|12)-(0?[1-9]|[12][0-9]|3[01])|(0?4|0?6|0?9|11)-(0?[1-9]|[12][0-9]|30)|(0?2)-(0?[1-9]|[12][0-9]))$",
            "MM月dd日,M月d日": r"((0?1|0?3|0?5|0?7|0?8|10|12)月(0?[1-9]|[12][0-9]|3[01])日|(0?4|0?6|0?9|11)月(0?[1-9]|[12][0-9]|30)日|(0?2)月(0?[1-9]|[12][0-9])日)$",
            "MM月dd日(ddd),M月d日(ddd)": r"((0?1|0?3|0?5|0?7|0?8|10|12)月(0?[1-9]|[12][0-9]|3[01])日|(0?4|0?6|0?9|11)月(0?[1-9]|[12][0-9]|30)日|(0?2)月(0?[1-9]|[12][0-9])日)\([日月火水木金土]\)$",
        }

        dateformat = ["%Y%m%d", "%Y/%m/%d", "%Y-%m-%d", "%Y年%m月%d日", "%Y年%m月%d日(%a)", "%m%d", "%m/%d", "%m-%d", "%m月%d日", "%m月%d日(%a)"]
        weekday_str = ["(日)","(月)","(火)","(水)","(木)","(金)","(土)"]

        for i, key  in enumerate(pattern_dic):
            pattern = pattern_dic[key]
            m = re.match(pattern, date_str)
            if m == None:
                continue
            else:
                i = (i + 1) % 10
                result_str = datetime.datetime.now().strftime(dateformat[i])
                if i == 4 or i == 9:
                    weekday_index = int(datetime.datetime.now().strftime("%w"))
                    result_str = datetime.datetime.now().strftime(dateformat[i-1]) + weekday_str[weekday_index]
    return result_str