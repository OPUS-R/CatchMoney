import nfc  #NFCライブラリ
from typing import cast
import tkinter as tk    #GUI制作ライブラリ
from tkinter import StringVar
import threading
import gspread
from oauth2client.service_account import ServiceAccountCredentials  #GCPO認証ライブラリ
from datetime import datetime
import serial
import discord  #discordbotライブラリ
from discord.ext import commands
import asyncio
import json
import os

##################################################
#別の組織で使う際は以下の変更が必要                    #
#①29行目グーグル認証キーの変更                       #
#②34行目以下数行で各シートの変更                     #
#③74行目ディスコードtoken変更                      #
#④151,152行目NFCのシステム、サービスコード変更        #
#⑤160行目学籍番号を出すために前後何文字を消すか変更     #
#⑥185行目必要に応じてタブリストに載せるシート名変更     #
##################################################


# Google Sheets 認証
#使用シートマスターが違う場合、キーを変える必要あり。
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]  #いじる必要無し
creds = ServiceAccountCredentials.from_json_keyfile_name("GCOA.json", scope)    #同ファイル内の場合は.json。別の場所の場合パス入力。
client = gspread.authorize(creds)

# スプレッドシート指定
#認証キーjson内のclient_emailのアドレスをシートの共有に編集者で追加するのを忘れない。
main_spreadsheet = client.open("名簿DB")  # NFCで読み取った学籍番号と名前を突合させるスプレッドシート
worksheet = main_spreadsheet.sheet1  # 最初のシートを使用

debt_spreadsheet = client.open("2025年度会計書類")  # 滞納関連のスプレッドシート/集金計測用シート。滞納と名前に就くものをタブに表示
master_spreadsheet = client.open("マスターDB")  #管理者権限管理用シート。学籍番号と権限を突合させる
master_worksheet = master_spreadsheet.sheet1    #最初のシートを使用

# ログファイルのパス
LOG_FILE = "log.txt"
LOG_MONEY="money.txt"
# シリアル通信設定（COMポートは適時調整）
#SERIAL_PORT = "COM4"
#BAUD_RATE = 9600
#ser = serial.Serial(SERIAL_PORT, BAUD_RATE)

#discord_channeljson
CHANNEL_CONFIG_PATH = "discord_channel.json"

def load_active_channel_id():
    global active_channel_id
    if os.path.exists(CHANNEL_CONFIG_PATH):
        try:
            with open(CHANNEL_CONFIG_PATH, "r", encoding="utf-8") as f:
                config = json.load(f)
                active_channel_id = config.get("active_channel_id")
                if active_channel_id:
                    log_message(f"保存された通知チャンネルIDを読み込み: {active_channel_id}")
        except Exception as e:
            log_message(f"通知チャンネルIDの読み込みエラー: {e}")


# GUIのセットアップ
root = tk.Tk()
root.title("CMbyOPUS")    #上に表示させる名前
root.geometry("400x900")    #GUIの大きさ。モニターに合わせる

# ラベルと入力用変数
student_num_var = StringVar()   #枠①の値のグローバル,学籍番号格納
student_num_var.set("読み取り待機中...")#枠①初期設定

result_var = StringVar()    #枠②、学籍番号と突合させた名前の格納
result_var.set("結果待機中...")

input_value_var = StringVar()   #枠④、支払い値段の格納
input_value_var.set("0")  # デフォルトは0

serial_value_var = StringVar()  #枠⑤、支払った金額(シリアル通信の値の乗算)の格納
serial_value_var.set("0")  # デフォルトは0

selected_tab_var = StringVar()  #選択した滞納リスト
selected_tab_var.set("タブ未選択")  # 初期値

log_text = tk.Text(root, font=("Arial", 12), height=10, wrap=tk.WORD)

is_permitted = False  # 管理者確認用フラグ

#Discord token設定
TOKEN = json.load(open("discord.json"))["discord_token"]    #同ファイル内の場合は.json。別の場所の場合パス入力。

#Discordのインテント有効化
intents = discord.Intents.default()
intents.messages = True
intents.guilds = True
intents.message_content = True
bot = commands.Bot(command_prefix=";", intents=intents)
active_channel_id = None


# ログ表示と保存
def log_message(message):
    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    log_entry = f"{timestamp} {message}\n"

    log_text.insert(tk.END, log_entry)
    log_text.see(tk.END)

    with open(LOG_FILE, "a", encoding="utf-8") as file:
        file.write(log_entry)


def money_message(message):
    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    money_entry = f"{timestamp} {message}\n"

    with open(LOG_MONEY, "a", encoding="utf-8") as file:
        file.write(money_entry)

# GUI(ボタン除く)
frame_label = tk.Label(root, text="枠① (学籍番号):", font=("Arial", 12))
frame_label.pack(pady=5)

frame_display = tk.Label(root, textvariable=student_num_var, font=("Arial", 14), bg="white", relief="solid", width=20, height=2)
frame_display.pack(pady=5)

result_label = tk.Label(root, text="枠② (照合結果):", font=("Arial", 12))
result_label.pack(pady=5)

result_display = tk.Label(root, textvariable=result_var, font=("Arial", 14), bg="white", relief="solid", width=20, height=2)
result_display.pack(pady=5)

input_label = tk.Label(root, text="枠④ 徴収金額:", font=("Arial", 12))
input_label.pack(pady=5)

input_entry = tk.Entry(root, textvariable=input_value_var, font=("Arial", 14), width=20)
input_entry.pack(pady=5)

serial_label = tk.Label(root,text="枠⑤ 支払い金額:", font=("Arial", 12))
serial_label.pack(pady=5)

#serial_display = tk.Entry(root, textvariable=serial_value_var, font=("Arial", 14), bg="white", relief="solid", width=20, height=2)
serial_display = tk.Entry(root, textvariable=serial_value_var, font=("Arial", 14), bg="white", relief="solid", width=20)
serial_display.pack(pady=5)

dropdown_label = tk.Label(root, text="タブ選択:", font=("Arial", 12))
dropdown_label.pack(pady=5)

dropdown_menu = tk.OptionMenu(root, selected_tab_var, "タブ未選択")
dropdown_menu.pack(pady=5)



# シリアル通信の値を代入
def read_serial():
    while True:
        try:
            if ser.in_waiting > 0:
                line = ser.readline().decode("utf-8").strip()
                current_value = int(serial_value_var.get())
                serial_value_var.set(str(current_value + int(line)))
                log_message(f"シリアルデータ受信: {line}")
        except Exception as e:
            log_message(f"シリアル通信エラー: {str(e)}")

# NFCの処理
nfc_active = threading.Event()
nfc_active.set()

def on_connect(tag):    #学生証読み取り
    global nfc_active
    if not nfc_active.is_set():
        return

    sys_code = 0xFE00   #学籍番号が格納されているシステムコード。学校毎にdumpして確認する必要あり
    service_code = 0x1A8B   #学籍番号が格納されているサービスコード。学校毎にdumpして確認する必要あり
    idm, pmm = tag.polling(system_code=sys_code)
    tag.idm, tag.pmm, tag.sys = idm, pmm, sys_code
    sc = nfc.tag.tt3.ServiceCode(service_code >> 6, service_code & 0x3F)

    bc = nfc.tag.tt3.BlockCode(0, service=0)
    student_num = cast(bytearray, tag.read_without_encryption([sc], [bc]))
    student_num = student_num.decode("shift_jis")
    if len(student_num) > 8:
        student_num = student_num[2:-6] #学校毎に形式の確認必要あり。(前後何文字削れば良いのか)
    student_num_var.set(student_num)    #抜き出した学籍番号をグローバルに格納
    log_message(f"NFC 読み取り成功: {student_num}")

    try:#枠①に生徒番号を記入
        cell = worksheet.find(student_num)
        result = worksheet.cell(cell.row, cell.col + 1).value
        result_var.set(result)
    except gspread.exceptions.CellNotFound:
        result_var.set("見つかりません")
        log_message("該当データが見つかりません")

    selected_tab = selected_tab_var.get()
    if selected_tab != "タブ未選択":
        try:
            tab_sheet = debt_spreadsheet.worksheet(selected_tab)
            cell = tab_sheet.find(result_var.get())
            next_cell_value = tab_sheet.cell(cell.row, cell.col + 4).value  # 対象列（必要なら +1 や +2 に調整）
            input_value_var.set(next_cell_value)
            serial_value_var.set(next_cell_value)
            log_message(f"NFC読み込み: {selected_tab} の {result_var.get()} の金額を反映: {next_cell_value}")
        except Exception as e:
            log_message(f"NFC読み込み金額取得エラー: {e}")

    nfc_active.clear()

def nfc_reader_loop():
    while True:
        try:
            with nfc.ContactlessFrontend("usb") as clf:
                log_message("NFCデバイスに接続しました")
                while True:
                    if nfc_active.is_set():
                        clf.connect(rdwr={"on-connect": on_connect})
        except Exception as e:
            log_message(f"NFC接続エラー（再試行します）: {e}")
            import time
            time.sleep(30)  # 少し待って再接続

#def nfc_reader_loop():
    #with nfc.ContactlessFrontend("usb") as clf:
        #while True:
           # if nfc_active.is_set():
               # clf.connect(rdwr={"on-connect": on_connect})

# ドロップダウンメニュー更新
def update_dropdown_menu():
    try:
        sheets = debt_spreadsheet.worksheets()
        debt_tabs = [sheet.title for sheet in sheets if "滞納" in sheet.title]    #滞納と名前のつくシートを認識。必要に応じて変更
        menu = dropdown_menu["menu"]
        menu.delete(0, "end")
        for tab in debt_tabs:
            menu.add_command(label=tab, command=lambda value=tab: on_tab_selected(value))
        if debt_tabs:
            selected_tab_var.set(debt_tabs[0])
            on_tab_selected(debt_tabs[0])
        else:
            selected_tab_var.set("タブ未選択")
        log_message("タブリスト更新成功")
    except Exception as e:
        log_message(f"タブリスト更新失敗: {str(e)}")

# タブ選択時の処理
def on_tab_selected(tab_name):
    selected_tab_var.set(tab_name)
    log_message(f"タブ選択: {tab_name}")

    try:
        worksheet = debt_spreadsheet.worksheet(tab_name)
        search_key = result_var.get()

        # 枠②の検索キーが設定されている場合のみ処理
        if search_key != "結果待機中...":
            try:
                cell = worksheet.find(search_key)
                next_cell_value = worksheet.cell(cell.row, cell.col + 4).value  # 一つ右のセルの値
                input_value_var.set(next_cell_value)  # 枠④に設定
                serial_value_var.set(next_cell_value)
                log_message(f"{tab_name} の '{search_key}' の右隣の値を設定: {next_cell_value}")
            except gspread.exceptions.CellNotFound:
                log_message(f"エラー: {search_key} が {tab_name} 内で見つかりません")
                input_value_var.set("0")
        else:
            log_message("枠②に有効な検索キーが設定されていません")
            input_value_var.set("0")
    except Exception as e:
        log_message(f"エラー: タブ選択処理中に問題が発生しました - {str(e)}")   #セル更新時のエラーは無視して大丈夫
        input_value_var.set("0")

# result_varの監視処理
def result_var_updated(*args):
    search_key = result_var.get()
    selected_tab = selected_tab_var.get()

    if selected_tab == "徴収表未選択":
        log_message("エラー: タブが選択されていません")
        return

    try:
        worksheet = debt_spreadsheet.worksheet(selected_tab)
        if search_key not in ["", "参照待機中...", "見つかりません"]:
            try:
                cell = worksheet.find(search_key)
                next_cell_value = worksheet.cell(cell.row, cell.col + 1).value  # 一つ右のセルの値
                input_value_var.set(next_cell_value)  # 枠④に設定
                output_value_var.set(next_cell_value)
                log_message(f"{selected_tab} の '{search_key}' の右隣の値を設定: {next_cell_value}")
            except gspread.exceptions.CellNotFound:
                log_message(f"エラー: {search_key} が {selected_tab} 内で見つかりません")
                input_value_var.set("0")
        else:
            log_message("枠②に有効な検索キーが設定されていません")
            input_value_var.set("0")
    except Exception as e:
        log_message(f"エラー: タブ検索処理中に問題が発生しました - {str(e)}")
        input_value_var.set("0")

# result_varにトリガーを設定
result_var.trace("w", result_var_updated)


# セルを更新:処理
def update_cell_in_selected_tab():
    try:
        selected_tab = selected_tab_var.get()
        search_key = result_var.get()
        serial_value = serial_value_var.get()
        student_num = student_num_var.get()
        timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")

        if selected_tab == "タブ未選択":
            log_message("エラー: タブが選択されていません")
            result_var.set("タブ未選択")
            return

        if not serial_value.isdigit():
            log_message("エラー: シリアル値が無効です")
            result_var.set("無効な入力")
            return

        if search_key == "結果待機中...":
            log_message("エラー: 検索対象が設定されていません")
            result_var.set("無効な検索対象")
            return

        worksheet = debt_spreadsheet.worksheet(selected_tab)
        try:
            cell = worksheet.find(search_key)
        except gspread.exceptions.CellNotFound:
            log_message(f"エラー: 該当文字列 '{search_key}' が見つかりません")
            result_var.set("該当なし")
            return

        worksheet.update_cell(cell.row, cell.col + 2, serial_value)
        result_var.set("更新成功")
        log_message(f"{selected_tab} のセルを更新しました: {search_key} -> {serial_value}")
        money_message(f"{selected_tab} のセルを更新しました: {search_key} -> {serial_value}")
        asyncio.run_coroutine_threadsafe(send_discord_message(f"{timestamp}{student_num}{search_key}が{selected_tab}{serial_value}円払った"), bot.loop)
        reset_detection()

    except Exception as e:
        log_message(f"エラー: 更新処理中に問題が発生しました - {str(e)}")
        result_var.set("更新失敗")

# NFCをリセット
def reset_detection():
    nfc_active.set()
    student_num_var.set("読み取り待機中...")
    result_var.set("結果待機中...")
    serial_value_var.set("0")
    log_message("NFC 待機状態にリセットしました")

#管理者権限ID判別
def check_permission():
    global is_permitted
    student_num = student_num_var.get()

    # 前回の student_num を記録し、変更があったときのみログを出力
    if hasattr(check_permission, "last_student_num") and check_permission.last_student_num == student_num:
        return  # 同一人物ならスキップ

    check_permission.last_student_num = student_num  # student_num を記録

    try:
        # student_num のセルを検索
        cell = master_worksheet.find(student_num)

        # cell が None でないか確認:エラー回避
        if cell is not None:
            # 右側のセルの値を取得
            master_status = master_worksheet.cell(cell.row, cell.col + 1).value

            # "管理者" なら is_permitted を True にする
            if master_status == "管理者":
                is_permitted = True
                log_message(f"管理者がログイン: {student_num}")
            else:   #リストにあるが管理者でない場合、別権限者としてログイン(今は仮で利用者)
                is_permitted = False
                log_message(f"利用者がログイン: {student_num} (ステータス: {master_status})")
        else:   #リストに名前がない。利用者のログインを記録
            is_permitted = False
            log_message(f"利用者がログイン '{student_num}' リスト無")

    except gspread.exceptions.CellNotFound:#空白だった場合の処理。基本見ない
        is_permitted = False
        log_message(f"管理者ではありません（検索結果なし: {student_num}）")

    update_ui_state()   #管理者の判別に基づきボタンの操作許可の処理を動かす
    root.after(5000, check_permission)  #msに1貝管理者判別の処理を実行。

#管理者権限持ちに各ボタンの使用を許可する
def update_ui_state():
    state = tk.NORMAL if is_permitted else tk.DISABLED
    #dropdown_menu.config(state=state)
    #debt_tab_button.config(state=state)
    #update_button.config(state=state)


# ボタン配置
reset_button = tk.Button(root, text="検出再開 (リセット)", font=("Arial", 12), command=reset_detection)
reset_button.pack(pady=10)



debt_tab_button = tk.Button(root, text="タブ更新", font=("Arial", 12), command=update_dropdown_menu)
debt_tab_button.pack(pady=10)


update_button = tk.Button(root, text="支払い", font=("Arial", 12), command=update_cell_in_selected_tab)
update_button.pack(pady=10)

autentic_button = tk.Button(root, text="管理者照合", font=("Arial", 12), command=check_permission)
autentic_button.pack(pady=10,padx=10)


log_label = tk.Label(root, text="ログ:", font=("Arial", 12))
log_label.pack(pady=5)

log_text.pack(pady=5)


#####以下discordのnotify botに関する関数#####


@bot.event  #notify botの起動確認用,非同期
async def on_ready():
    print(f"notify bot起動: {bot.user}")
    asyncio.run_coroutine_threadsafe(send_discord_message(f"discord bot 接続成功 "), bot.loop)

@bot.command()
async def cm(ctx, channel_name: str):
    global active_channel_id
    target_channel = discord.utils.get(ctx.guild.text_channels, name=channel_name)

    if target_channel:
        active_channel_id = target_channel.id
        await ctx.send(f"✅ {target_channel.mention} 通知チャンネルを設定しました")

        # 保存
        try:
            with open(CHANNEL_CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump({"active_channel_id": active_channel_id}, f)
                log_message(f"通知チャンネルID {active_channel_id} を保存しました")
        except Exception as e:
            log_message(f"通知チャンネルIDの保存エラー: {e}")

    else:
        await ctx.send(f"❌ チャンネル `{channel_name}` が見つかりません。")


@bot.event  #メッセージ送信の際に、;cm <チャンネル名>で指定したチャンネル以外に通知を送らないようにする関数,非同期処理
async def on_message(message):
    global active_channel_id

    # BOT自身のメッセージを無視する
    if message.author.bot:
        return

    # チャンネルが指定されてる場合のみ反応
    if active_channel_id is not None and message.channel.id != active_channel_id:
        return  # 指定チャンネル以外でのメッセージの無視

    # 通常のコマンド処理
    await bot.process_commands(message)

#BOT起動

def start_discord_bot():
    bot.run(TOKEN)

# Discordメッセージ送信非同期処理
async def send_discord_message(message):
    global active_channel_id
    if active_channel_id:
        channel = bot.get_channel(active_channel_id)
        if channel:
            await channel.send(message)

#discord bot 別スレッド起動main.loopとの回避
load_active_channel_id()
bot_thread = threading.Thread(target=start_discord_bot, daemon=True)
bot_thread.start()


#####discord関連ここまで#####


# 起動時にタブを更新
update_dropdown_menu()

# NFCスレッド開始
nfc_thread = threading.Thread(target=nfc_reader_loop, daemon=True)
nfc_thread.start()

# シリアル通信スレッド開始
#serial_thread = threading.Thread(target=read_serial, daemon=True)
#serial_thread.start()

#管理者権限判別起動
check_permission()

# GUIメインループ
root.mainloop()