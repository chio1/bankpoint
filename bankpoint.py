import discord
import asyncio

client = discord.Client()



@client.event
async def on_ready():
    print("로그인 합니다.")
    print(client.user.name)
    print(client.user.id)
    print('------')


client.run('NTkyNjc1NjA0MzExMDQ4MjAy.XRD8wg.fZFG5TlXSOsleID_2Ax12GTdR8M')

