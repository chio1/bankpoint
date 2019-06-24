import discord
import asyncio
import os
import openpyxl

client = discord.Client()

@client.event
async def on_ready():
    print('온')
    print('------')
    game = discord.Game("!포인트확인 Made in Chio")
    await client.change_presence(status=discord.Status.online , activity=game)



@client.event
async def on_message(message):
    if message.content.startswith("!포인트주기"):
        if str(message.content[7:12]) == "":
            await message.channel.send("<명령어> !포인트주기 (고유번호)")
        else:
            author = message.content.startswith(str(message.content[7:12]))
            file = openpyxl.load_workbook("포인트.xlsx")
            sheet = file.active
            i = 1
            while True:
                if sheet["A" + str(i)].value == str(message.content[7:12]):
                    sheet["B" + str(i)].value = int(sheet["B" + str(i)].value) + 1
                    file.save("포인트.xlsx")
                    await message.channel.send("축하드립니다 진급포인트 1획득 ")
                    await message.channel.send(str(message.content[6:12]) +"님의 누적 포인트 : " + str(sheet["B" + str(i)].value))

                    break
                if sheet["A" + str(i)].value == None:
                    sheet["A" + str(i)].value = str(message.content[7:12])
                    sheet["B" + str(i)].value = 1
                    file.save("포인트.xlsx")
                    await message.channel.send("축하드립니다 진급포인트 1획득 ")
                    await message.channel.send(str(message.content[7:12]) +"님의 누적 포인트 : " + str(sheet["B" + str(i)].value))

                    break
                i += 1
    if message.content.startswith("!포인트확인"):
        if str(message.content[7:12]) == "":
            await message.channel.send("<명령어> !포인트확인 (고유번호)")
        else:
            author = message.content.startswith(str(message.content[6712]))
            file = openpyxl.load_workbook("포인트.xlsx")
            sheet = file.active
            i = 1
            while True:
                if sheet["A" + str(i)].value == str(message.content[7:12]):
                    await message.channel.send(str(message.content[7:12]) +"님의 누적 포인트 : " + str(sheet["B" + str(i)].value))
                    break
                else:
                    await message.channel.send(str(message.content[7:12]) + "님의 누적 포인트 : 0 " )
                    break

access_token = os.environ["BOT_TOKEN"]
client.run(access_token)

