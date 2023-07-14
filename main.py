import discord
import requests
import win32com.client
from discord.player import FFmpegPCMAudio
import ffmpeg
import os
from pyowm.owm import OWM
from pyowm.utils import formatting
from pyowm.utils import timestamps
from pyowm.utils.config import get_default_config

from pandas_datareader.stooq import StooqDailyReader as sdr

import wikipedia

from datetime import timezone, timedelta, datetime
import pytz

TOKEN = "OTA4NTMxNTE0OTczNzgyMDc5.G2CcSL.jm_6ZJ09PtqRXCAmOsjSG5De5NWqSSB27s70Yk"

# wikipedia Library setting
wikipedia.set_lang("ja")

# Openwether API setting and PyOWM config setting
API_KEY = "47b6eef3d271defa1dc8150c064a01bb"
config_dict = get_default_config()
config_dict["language"] = "ja"
owm = OWM(API_KEY, config_dict)
mgr = owm.weather_manager()

# Rakuten API setting
RAKU_TRAVEL_API_URL = "https://app.rakuten.co.jp/services/api/Travel/KeywordHotelSearch/20170426?"
RAKU_API_ID = "1001723775230610155"
RAKU_AFFILIATE_ID = '33389144.28ae5147.33389145.82470423'

# timezone setting
jst = pytz.timezone('Asia/Tokyo')

# Client setting
client = discord.Client(intents=discord.Intents.all())

# HotPepper API Setting
HP_URL = "http://webservice.recruit.co.jp/hotpepper/gourmet/v1/"
URLHP = "https://www.hotpepper.jp"
HOT_PEPPER_API_KEY = "c53e1ecf09bb3a69"

# CevioAI API Setting(Cast Toggle comment out) & (now stoping service voice talk bot)
cevio = win32com.client.Dispatch("CeVIO.Talk.RemoteService2.ServiceControl2V40")
cevio.StartHost(False)
talker = win32com.client.Dispatch("CeVIO.Talk.RemoteService2.Talker2V40")
# talker.Cast = "IA"
talker.Cast = "小春六花"


@client.event
async def on_ready():
    print('ログインしました')


@client.event
async def on_message(message):
    if message.author.bot:
        return

    if message.content == "にゃーん":
        await message.channel.send('にゃーん')

    if message.content == "はにゃーん":
        await message.channel.send("は？？")

    if message.content.startswith("/grm"):
        return

    if message.content.startswith("/weather"):
        where = message.content.split()
        place = where[1]
        observation = mgr.weather_at_place(place)
        w = observation.weather
        time = formatting.to_date(w.ref_time).astimezone(jst)
        print(format(formatting.to_date(w.ref_time).replace(tzinfo=jst)))
        print(format(w.detailed_status))
        print(format(w.temperature("celsius")))
        print(format(w.barometric_pressure()))
        print(format(w.rain))
        print(format(w.snow))
        value = ("計測時間 : " + time.strftime('%Y年%m月%d日 %H:%M:%S') + "\n天気 : " + format(
            w.detailed_status) + "\n気温 : " + format(w.temperature("celsius")["temp"]) + "℃\n湿度 : " + format(
            w.humidity) + "%")
        if "1h" in w.rain:
            value += "\n雨量 : " + format(w.rain) + "mm/時"
        else:
            value += "\n雨量 : 現在雨は降っていません。"
        if "1h" in w.snow:
            value += "\n積雪 : " + format(w.snow) + "mm/時"
        await message.channel.send(value)

    if message.content.startswith("/forecast"):
        where = message.content.split()
        place = where[1]
        observation = mgr.forecast_at_place(place, "24h")
        tomorrow = timestamps.tomorrow()
        w = observation.get_weather_at(tomorrow)
        await message.channel.send(w)

    if message.content.startswith("/wiki"):
        keyword = message.content.split()
        key = keyword[1]
        try:
            wp = wikipedia.page(key)
            if len(wp.summary) >= 1900:
                await message.channel.send("ごめんなさい文字数オーバーでした...")
            else:
                await message.channel.send(wp.title + " ：\n " + wp.summary)
        except wikipedia.exceptions.DisambiguationError as e:
            await message.channel.send(e)

    if message.content.startswith("/hp"):
        s = message.content.strip("/hp ")
        keyword = s.replace("　", " ")
        body = {
            'key': HOT_PEPPER_API_KEY,
            'keyword': keyword,
            'format': 'json',
            'count': 1
        }
        res = requests.get(HP_URL, body)
        data = res.json()
        print(data)
        stores = data['results']['shop']
        hit = int(data['results']['results_returned'])
        hit2 = int(data['results']['results_available'])
        print(hit2)
        for stores_name in stores:
            name = stores_name['name']
            urls = stores_name['urls']['pc']
            catch = stores_name['catch']
        if hit == 0:
            await message.channel.send("私には見つけられなかったので" + URLHP + " でどうぞ")
        else:
            await message.channel.send("# " + name + "\n" + catch + "\n" + urls)

    if message.content.startswith("/rakutra"):
        s = message.content.strip("/rakutra ")
        keyword = s.replace("　", " ")
        body = {
            'applicationId': RAKU_API_ID,
            'affiliateId': RAKU_AFFILIATE_ID,
            'format': 'json',
            'carrier': 0,
            'page': 1,
            'hits': 1,
            'keyword': keyword,
            'formatVersion': 2
        }
        res = requests.get(RAKU_TRAVEL_API_URL, body)
        data = res.json()
        hotel_policy = data['hotels']['hotelBasicInfo']
        hotel = data['hotels']['hotelPolicyInfo']
        print(data)
        hit = int(data['pagingInfo']['recordCount'])
        print(hit)
        for hotels_name in hotel:
            hname = hotels_name['hotelName']
            url = hotels_name['hotelInformationUrl']
            summary = hotels_name['hotelSpecial']
            min_Charge = hotels_name['hotelMinCharge']
            address = hotels_name['address1'] + hotels_name['address2']
            parking = hotels_name['parkingInformation']
            station = hotels_name['nearestStation']
        for hotels_pi in hotel_policy:
            card = hotels_pi['availableCreditCard']
        if hit == 0:
            await message.channel.send("すいません、見つかりませんでした")
        elif card == None:
            await message.channel.send(
                "# " + hname + "\n" + summary + "\n 最低金額：" + min_Charge + "\n" + address + "\n 駐車場：" + parking + "\n 最寄り駅：" + station + "\n " + "\n" + url)
        else:
            await message.channel.send(
                "# " + hname + "\n" + summary + "\n 最低金額：" + min_Charge + "\n" + address + "\n 駐車場：" + parking + "\n 最寄り駅：" + station + "\n 利用可能カード：" + card + "\n" + url)

    # if message.content.startswith("/trade"):
    # keyword = message.content.split()
    # Number = []
    # Number.append(str(keyword[1] + '.T'))
    # print(Number)
    # date:str = keyword[2]
    # df = web.DataReader('7203.JP', 'stooq', start='2023-06-17', end='2023-06-18')
    # await message.chanel.send(df)
    # stooq = sdr('AAPL.US',start=datetime(2023,6,10),end=datetime(2023,6,11))
    # data = stooq.read()
    # print(data)

    if message.content == "/join":
        if message.author.voice is None:
            await message.channel.send("あなたはボイスチャンネルに接続していません")
            return

        await message.author.voice.channel.connect()
        await message.channel.send("接続しました")

    if message.content == "/leave":
        if message.guild.voice_client is None:
            await message.channel.send("接続していません")
            return

        await message.guild.voice_client.disconnect()
        await message.channel.send("切断しました")

    if message.author.bot:
        return
    elif message.content != None:  # CevioAI Command
        print("よばれました")
        if message.channel.id == 857953614387609627 or message.channel.id == 898200509687681074:
            print("よばれました2")
            if message.guild.voice_client is None:
                return

            elif message.content.startswith("```"):
                state = talker.OutputWaveToFile("コードブロックです", "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            elif message.content.startswith("___"):
                msg = message.content
                msg = msg[4:-4]
                state = talker.OutputWaveToFile(msg, "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            elif message.content.startswith("**"):
                msg = message.content
                msg = msg[3:-3]
                state = talker.OutputWaveToFile(msg, "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            elif message.content.startswith("*"):
                msg = message.content
                msg = msg[2:-2]
                state = talker.OutputWaveToFile(msg, "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            elif message.content.startswith("_"):
                msg = message.content
                msg = msg[4:-4]
                state = talker.OutputWaveToFile(msg, "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            elif message.content.startswith("__"):
                msg = message.content
                msg = msg[3:-3]
                state = talker.OutputWaveToFile(msg, "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            elif message.content.startswith("||"):
                state = talker.OutputWaveToFile("ネタバレ防止", "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            elif message.content.startswith("/spoiler"):
                state = talker.OutputWaveToFile("ネタバレ防止", "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            elif message.content.startswith("~~"):
                msg = message.content
                msg = msg[3:-3]
                state = talker.OutputWaveToFile(msg, "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            elif message.content.startswith("`"):
                msg = message.content
                msg = msg[2:-2]
                state = talker.OutputWaveToFile(msg, "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            elif message.content.startswith(">"):
                state = talker.OutputWaveToFile("msg", "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            elif message.content.startswith("http"):
                state = talker.OutputWaveToFile("URL省略", "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")

            else:
                state = talker.OutputWaveToFile(message.content, "C:\AudioOutFile\Audio.wav")
                if state == True:
                    print("True")
                elif state == False:
                    print("False")
                await message.guild.voice_client.play(discord.FFmpegPCMAudio("C:\AudioOutFile\Audio.wav", ))
                os.remove("C:\AudioOutFile\Audio.wav")


client.run(TOKEN)
