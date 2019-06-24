import discord
import asyncio
import os

client = discord.Client()



@client.event
async def on_ready():
    print("로그인 합니다.")
    print(client.user.name)
    print(client.user.id)
    print('------')

access_token = os.environ["BOT_TOKEN"]
client.run(access_token)

