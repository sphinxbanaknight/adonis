import discord
import random
import os
import json
import gspread
import pprint
from oauth2client import file as oauth_file, client, tools
from apiclient.discovery import build
from httplib2 import Http
import time
import datetime
import pytz
import asyncio


from pytz import timezone
from datetime import datetime, timedelta

from oauth2client.service_account import ServiceAccountCredentials
from discord.ext import commands, tasks

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

basedir = os.path.abspath(os.path.dirname(__file__))
data_json = basedir+'/cogs/client_secret.json'

creds = ServiceAccountCredentials.from_json_keyfile_name(data_json, scope)
gc = gspread.authorize(creds)

takte = gc.open('Tempo')
rostersheet = takte.worksheet('WoE Roster')
fullofsheet = takte.worksheet('Full IGNs')
datasheet = takte.worksheet('Data')

################ Channel, Server, and User IDs ###########################
sphinx_id = 108381986166431744
ardi_id = 248681868193562624
kriss_id = 694307907835134022
ken_id = 158345623509139456
jude_id = 693741143313088552
cell_id = 192286855025262592
glock_id = 706842108832776223
#servers = [401186250335322113, 691130488483741756, 800129405350707200]
sk_server = 401186250335322113
bk_server = 691130488483741756
c_server = 800129405350707200
servers = [sk_server, bk_server, c_server]

sk_bot = 401212001239564288
bk_bot = 691205255664500757
bk_ann = 695801936095740024 #BK #announcement
c_bot = 800129405350707200
botinit_id = [sk_bot, bk_bot, c_bot]
authorized_id = [sphinx_id, ardi_id, kriss_id, ken_id, jude_id, cell_id, glock_id]
dev_id = [sphinx_id]


################ Cell placements ################
guild_range = "B3:E99"
roster_range = "G3:I50"
matk_range = "L3:M14"
p1role_range = "P3:P14"
atk_range = "L17:M28"
p2role_range = "P17:P28"
p3_range = "L32:M43"
p3role_range = "P32:P43"

################ Parameters ################
# These flags are status trackers to avoid duplicate runs in the timed events
isarchived = False # archiving
isreminded1 = False # reminder wed
isreminded2 = False # reminder sat

# Default timed event runtime (format = "%H:%M:%A")
t_silk2_archive = '00:00:Sunday'
t_silk4_archive = '00:00:Monday'
tf_archive = [t_silk2_archive, t_silk4_archive]
tf_remind1 = [t_silk4_archive]
tf_remind2 = ['22:00:Wednesday', '22:00:Friday']
tf_reset   = ['00:05:Sunday', '00:05:Monday', '22:05:Wednesday']

# Responses
answer_timedevent_archive = ['archive', 'archiving']
answer_timedevent_remind1 = ['remind1']
answer_timedevent_remind2 = ['remind2']
answer_timedevent_reset = ['reset', 'refresh', 'init']

isremindenabled = False # configuration - turn on/off auto-reminder


################ Feedbacks ################
feedback_properplz = 'Please send a proper syntax: '
feedback_debug = '`[DEBUGINFO] `'
feedback_automsg = '`[Auto-generated message]` '
feedback_noangrypingplz = 'to avoid <:AngryPing:703193629489102888> from me, please register your attendance before coming Wed/Fri 10PM GMT+8 by typing `/att y/n, y/n` in <#691205255664500757>. Thank you!'


prefix = ["/"]
description = "A bot for sheet+discord linking/automation."
intents = discord.Intents.all()
client = commands.Bot(command_prefix = prefix, description = description, intents = intents)



client.remove_command('help')

def istimedeventformat(input):
    try:
        datetime.strptime(input, '%H:%M:%A')
        return True
    except ValueError:
        return False

def next_available_row(sheet, column):
    cols = sheet.range(3, column, 1000, column)
    return max([cell.row for cell in cols if cell.value]) + 1

def pinger():
    attlist = [item for item in rostersheet.col_values(7) if item and item != 'IGN' and item != 'Next WOE:']
    ignlist = [item for item in rostersheet.col_values(3) if item and item != 'IGN' and item != 'READ THE NOTES AT [README]']
    idlist = [item for item in fullofsheet.col_values(3) if item and item != 'UNIQUE:' and item != 'Discord Tag' and item != 'READ THE NOTES AT [README]']
    row = 3
    dsctag = []
    dscid = []
    
    for ign in ignlist:
        gottem = 0
        for att in attlist:
            if ign == att:
                ign = ""
                gottem = 1
                break
        if gottem == 0:
            try:
                dsctag.append(rostersheet.cell(row, 2).value)
            except Exception as e:
                print(f'Exception caught at dsctag: {e}')
        row += 1
        
    row = 4
    for idd in idlist:
        for dsc in dsctag:
            #print(f'{idd} to {dsc} has idd of {fullofsheet.cell(row, 2).value}')
            if idd == dsc:
                try:
                    dscid.append(fullofsheet.cell(row, 2).value)
                except Exception as e:
                    print(f'Exception caught at dsctag: {e}')
                break
        row += 1
    
    # Change to set for unique values only
    dscid_set = set(dscid)
    
    return dscid_set         

# get debugmode from Clears
def get_debugmode():
    clearscog = client.get_cog('Clears')
    debugger = clearscog.get_debugmode()
    return debugger


@client.event
async def on_ready():
    global isarchived
    global isreminded1
    global isreminded2
    global tf_archive
    global tf_remind1
    global tf_remind2
    global tf_reset
    #global cresence

    for server in client.guilds:
        if server.id == sk_server:
            sphinx = server
        if server.id == c_server:
            cresence = server
            break
        if server.id == bk_server:
            burger = server

    for channel in sphinx.channels:
        if channel.id == sk_bot:
            botinitsk = channel
            break

    for channel in burger.channels:
        if channel.id == bk_bot:
            botinitbk = channel
        # elif channel.id == bk_ann:
        #     botinitbkann = channel

    print('Bot is online.')
    await client.change_presence(status=discord.Status.dnd, activity=discord.Game('Getting scolded by Cai'))


    shit = gc.open('Tempo')
    try:
        wsheet = shit.worksheet('WoE Roster Archive')
    except gspread.exceptions.WorksheetNotFound:
        await botinitsk.send(f'Could not find 5th sheet in our GSheets, creating one now.')
        #await botinitbk.send(f'Could not find 5th sheet in our GSheets, creating one now.')
        await botinitbk.send(f'Could not find 5th sheet in our GSheets, creating one now.')
        spreadsheet = gc.open('CRESENCE ROSTER')
        wsheet = spreadsheet.add_worksheet(title='WoE Roster Archive', rows = 1000, cols = 10)
        kekerino = wsheet.range("A1:J1000")
        for kek in kekerino:
            kek.value = ''
        wsheet.update_cells(kekerino, value_input_option='USER_ENTERED')

    print("Automated Clear Roster Begins!")
    format = "%H:%M:%A"

    ph_time = pytz.timezone('Asia/Manila')
    ph_time_unformated = datetime.now(ph_time)
    ph_time_formated = ph_time_unformated.strftime(format)


    data_pasted = [""]
    while True:
        debugger = get_debugmode()
        ph_time = pytz.timezone('Asia/Manila')
        ph_time_unformated = datetime.now(ph_time)
        ph_time_formated = ph_time_unformated.strftime(format)
        await asyncio.sleep(5)
            
        # if not isarchived and ph_time_formated in tf_archive:
        #     isarchived = True
        #     await botinitsk.send('```Automatically cleared the roster! Please use /att y/n, y/n again to register your attendance.```')
        #     await botinitsk.send('```An archive of the latest roster was saved in WoE Roster Archive Spreadsheet.```')
        #     await botinitbk.send('```Automatically cleared the roster! Please use /att y/n, y/n again to register your attendance.```')
        #     await botinitbk.send('```An archive of the latest roster was saved in WoE Roster Archive Spreadsheet.```')
        #     try:
        #         next_row = next_available_row(wsheet, 1)
        #     except ValueError as e:
        #         print(f'next_row returned {e}')
        #         next_row = 1
        #     copy_list2 = silk2.range("B4:D51")
        #     copy_list4 = silk4.range("B4:D51")
        #     paste_list = rostersheet.range(next_row, 1, next_row + 45, 3)
        #     count = 0
        #     newformat = "%B %Y"
        #     ph_time = pytz.timezone('Asia/Manila')
        #     ph_time_unformated = datetime.now(ph_time)
        #     ph_time_new_formated = ph_time_unformated.strftime(newformat)
        #     if ph_time_formated == t_silk2_archive:
        #         for copy in copy_list2:
        #             data_pasted.append(copy.value)
        #     elif ph_time_formated == t_silk4_archive:
        #         for copy in copy_list4:
        #             data_pasted.append(copy.value)
        #     for paste in paste_list:
        #         if count == 0:
        #             if ph_time_formated == t_silk2_archive:
        #                 d = ph_time_unformated.strftime("%d")
        #                 d = int(d)
        #                 d -= 1
        #                 d = str(d)
        #                 paste.value = f'{d} {ph_time_new_formated} PM SAT WOE'
        #                 cell_list = silk2.range("B4:D50")
        #                 for cell in cell_list:
        #                     cell.value = ""
        #                 silk2.update_cells(cell_list, value_input_option='USER_ENTERED')
        #             elif ph_time_formated == t_silk4_archive:
        #                 d = ph_time_unformated.strftime("%d")
        #                 d = int(d)
        #                 d -= 1
        #                 d = str(d)
        #                 paste.value = f'{d} {ph_time_new_formated} PM SUN WOE'
        #                 cell_list = silk4.range("B4:D50")
        #                 for cell in cell_list:
        #                     cell.value = ""
        #                 silk4.update_cells(cell_list, value_input_option='USER_ENTERED')
        #         elif count == 1:
        #             paste.value = ""
        #         else:
        #             try:
        #                 paste.value = data_pasted[count - 2]
        #             except IndexError:
        #                 break
        #         count += 1
        #     wsheet.update_cells(paste_list, value_input_option='USER_ENTERED')
        #
        #     cell_list = rostersheet.range(roster_range)
        #
        #     for cell in cell_list:
        #         cell.value = ""
        #
        #     rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')
        #     continue
        # # Timed event [auto-reminder]: a soft reminder message into #announcement. Remove on next event
        if isremindenabled and not isreminded1 and ph_time_formated in tf_remind1:
            if debugger: await botinitsk.send(f'{feedback_debug} {ph_time_formated} Reminder1 isreminded1={isreminded1} START')
            isreminded1 = True
            try:
                # different reminder message depending on timing
                if ph_time_formated == t_silk4_archive:
                    msgstr = f'''{feedback_automsg} Hi folks, hope you enjoyed your weekend! For next week, {feedback_noangrypingplz}'''
                else:
                    att_igns = [item for item in rostersheet.col_values(7) if item and item != 'IGN' and item != 'Next WOE:']
                    nratt = len(att_igns)
                    msgstr = f'''{feedback_automsg}
Hi all,

Currently we have {nratt} members who have registered their attendance, great job!
For those who haven't: {feedback_noangrypingplz}'''
                if debugger: #send to test server if on debugmode
                    msg1 = await botinitsk.send(msgstr)
                else:
                    msg1 = await botinitbk.send(msgstr)
                datasheet.update_cell(2, 9, str(msg1.id) ) # save as string to avoid Excel nr truncation
                if debugger: await botinitsk.send(f'{feedback_debug} msg1 ID saved: `{msg1.id}`')
            except Exception as e:
                await botinitsk.send(f'Error: `{e}`')
            if debugger: await botinitsk.send(f'{feedback_debug} {ph_time_formated} Reminder1 isreminded1={isreminded1} END')
            continue
        # Timed event [auto-reminder]: @mention per player who enlisted but not yet confirmed attendance
        elif isremindenabled and not isreminded2 and ph_time_formated in tf_remind2:
            if debugger: await botinitsk.send(f'{feedback_debug} Angrypinger2 isreminded2={isreminded2} START')
            isreminded2 = True
            try:
                try: #msg1 may not be found
                    await msg1.delete()
                except NameError as e:
                    msg1 = await botinitsk.send(f'Error: I forgot the message! Attempting to fetch by message ID...')
                    try: #find value saved in sheet
                        msgid = datasheet.cell(2,9).value
                        if debugger:
                            msg1 = await botinitsk.fetch_message(msgid)
                        else:
                            try: # fetch from announcement or bot
                                msg1 = await botinitbk.fetch_message(msgid)
                            except Exception as e:
                                try:
                                    msg1 = await botinitbk.fetch_message(msgid)
                                except Exception as e:
                                    pass
                        await msg1.delete()
                    except Exception as e:
                        await botinitsk.send(f'Error: `{e}`. Unable to find and delete message based on `{msgid}`. Please manually delete it.')
                except Exception as e:
                    await botinitsk.send(f'Error: `{e}`. Skipping message deletion...')
                
                ping_tags = pinger()
                taglist = ''
                if debugger: #send to test server if on debugmode
                    for discordtag in ping_tags:
                        discordtag = random.choice(dev_id) # for testing purpose, use only the developers' id!
                        taglist += '<@' + str(discordtag) + '>, '
                    if taglist != '':
                        msg1 = await botinitsk.send(f'{feedback_debug} {feedback_automsg} Hi {taglist}you have not registered your attendance yet. <:peeposad:702156649992945674> Next time, {feedback_noangrypingplz}')
                else:
                    for discordtag in ping_tags:
                        taglist += '<@' + str(discordtag) + '>, '
                    if taglist != '':
                        msg1 = await botinitbk.send(f'{feedback_automsg} Hi {taglist}you have not registered your attendance yet. <:peeposad:702156649992945674> Next time, {feedback_noangrypingplz}')
                datasheet.update_cell(2, 9, str(msg1.id) ) # save as string to avoid Excel nr truncation
                if debugger: await botinitsk.send(f'{feedback_debug} msg1 ID saved: `{msg1.id}`')
            except Exception as e:
                await botinitsk.send(f'Error: `{e}`')
            if debugger: await botinitsk.send(f'{feedback_debug} {ph_time_unformated.strftime("%H:%M:%S.%f:%A")} Angrypinger2 isreminded2={isreminded2} END')
            continue            
        # Timed event status reset
        elif ph_time_formated in tf_reset:
            isarchived = False
            isreminded1 = False
            isreminded2 = False
            if debugger: await botinitsk.send(f'{feedback_debug} `[Timed event status reset] isarchived={isarchived} isreminded1={isreminded1} isreminded2={isreminded2}`')
            # Default timed event runtime (format = "%H:%M:%A")
            tf_archive = [t_silk2_archive, t_silk4_archive]
            tf_remind1 = [t_silk4_archive]
            tf_remind2 = ['22:00:Wednesday', '22:00:Friday']
            tf_reset   = ['00:05:Sunday', '00:05:Monday', '22:05:Wednesday']
            if debugger: await botinitsk.send(f'{feedback_debug} `[Timed event timers reset] tf_archive={tf_archive} tf_remind1={tf_remind1} tf_remind2={tf_remind2} tf_reset={tf_reset}`')
            continue

            


@client.event
async def on_member_join(member):
    try:
        role = discord.utils.get(member.guild.roles, id=703643328406880287)
    except AttributeError as e:
        print(f'autorole returned as {e}')
        role = discord.utils.get(member.guild.roles, name="new recruit")
    await member.add_roles(role)
        


# toggle isremindenabled
@client.command()
async def togglereminder(ctx):
    global isremindenabled
    channel = ctx.message.channel
    commander = ctx.author
    if channel.id in botinit_id:
        if commander.id in authorized_id:
            try:
                isremindenabled = not isremindenabled
            except Exception as e:
                await ctx.send(e)
            await ctx.send(f'`Auto-reminder Enabled = {isremindenabled}`')
        else:
            await ctx.send(f'*Nice try pleb.*')
    else:
        await ctx.send(f'Wrong channel! Please use #bot.')

# Force a timed event run
@client.command()
async def forcetimedevent(ctx, *, arguments):
    debugger = get_debugmode()
    channel = ctx.message.channel
    commander = ctx.author
    if channel.id in botinit_id:
        if commander.id in authorized_id:
            arglist = [x.strip() for x in arguments.split(',')]
            no_of_args = len(arglist)
            #0 = name of timed event
            #1 = time to run
            
            # reinitialise to reset previous forces
            global tf_archive
            global tf_remind1
            global tf_remind2
            global tf_reset
            tf_archive = [t_silk2_archive, t_silk4_archive]
            tf_remind1 = [t_silk4_archive]
            tf_remind2 = ['22:00:Wednesday', '22:00:Friday']
            tf_reset   = ['00:05:Sunday', '00:05:Monday', '22:05:Wednesday']

            if no_of_args == 2:
                try:
                    eventname = arglist[0].lower()
                    eventtime = arglist[1]
                    if debugger: await ctx.send(f'{feedback_debug} input: {eventname}, {eventtime}')
                    
                    if not istimedeventformat(eventtime):
                        await ctx.send(f'{feedback_properplz} Time format should be in H:M:Day, e.g. `/forcetimedevent, remind1, 22:00:Tuesday`')
                        return
                    
                    if eventname in answer_timedevent_archive:
                        global isarchived
                        isarchived = False
                        tf_archive.append(eventtime)
                        await ctx.send(f'Timed event archive added to also run at {eventtime}. All schedules: {tf_archive}')
                    elif eventname in answer_timedevent_remind1:
                        global isreminded1
                        isreminded1 = False
                        tf_remind1.append(eventtime)
                        await ctx.send(f'Timed event remind1 added to also run at {eventtime}. All schedules: {tf_remind1}')
                    elif eventname in answer_timedevent_remind2:
                        global isreminded2
                        isreminded2 = False
                        tf_remind2.append(eventtime)
                        await ctx.send(f'Timed event remind2 added to also run at {eventtime}. All schedules: {tf_remind2}')
                    elif eventname in answer_timedevent_reset:
                        tf_reset.append(eventtime)
                        await ctx.send(f'Timed event reset added to also run at {eventtime}. All schedules: {tf_reset}')
                    else:
                        await ctx.send(f'''{feedback_properplz} Name format should be one of the following:
`archive` = {tf_archive}
`remind1` = {tf_remind1}
`remind2` = {tf_remind2}
`reset  ` = {tf_reset}
e.g. `/forcetimedevent, remind1, 22:00:Tuesday`''')
                except Exception as e:
                    await ctx.send(f'Error: `{e}`')
            else:
                await ctx.send(f'{feedback_properplz} `/forcetimedevent, <name>, <time>`, e.g. `/forcetimedevent, remind1, 22:00:Tuesday`')
                return
        else:
            await ctx.send(f'*Nice try pleb.*')
    else:
        await ctx.send(f'Wrong channel! Please use #bot.')

# for testing purpose
@client.command()
async def jytest(ctx):
    channel = ctx.message.channel
    commander = ctx.author
    if not channel.id in botinit_id:
        await ctx.send(f'Wrong channel! Please use #bot.')
        return
    elif not commander.id in authorized_id:
        await ctx.send(f'*Nice try pleb.*')
        return
    
    try:
        await ctx.send(f'`jytest` start')
        global dev_id
        dev_id = [commander.id]
        await ctx.send(f'Overwrote dev_id to you. dev_id={dev_id}')
        
        
        global isarchived
        global isreminded1
        global isreminded2
        global isremindenabled
        await ctx.send(f'isarchived={isarchived}, isreminded1={isreminded1}, isreminded2={isreminded2}, isremindenabled={isremindenabled}')
                        
        await ctx.send(f'`jytest` end')
    except Exception as e:
        await ctx.send(f'Error: `{e}`')



@client.command()
async def load(ctx, extension):
    client.load_extension(f'cogs.{extension}')
    await ctx.send(f'Cog: {extension}.py loaded')


@client.command()
async def unload(ctx, extension):
    client.unload_extension(f'cogs.{extension}')
    await ctx.send(f'Cog: {extension}.py unloaded')


@client.command()
async def reload(ctx, extension):
    client.unload_extension(f'cogs.{extension}')
    client.load_extension(f'cogs.{extension}')
    await ctx.send(f'Cog: {extension}.py reloaded')


for filename in os.listdir('./cogs'):
    if filename.endswith('.py'):
        client.load_extension(f'cogs.{filename[:-3]}')

@client.event
async def on_command_error(ctx, error):
    if isinstance(error, commands.MissingRequiredArgument):
        await ctx.send('Please pass in all required arguments.')
    if isinstance(error, commands.CommandNotFound):
        await ctx.send('Invalid command used.')

client.run(os.environ['token'], reconnect = True)