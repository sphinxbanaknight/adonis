import discord
import random
import os
import json
import gspread
import pprint
#import models
import io
from oauth2client import file as oauth_file, client, tools
from apiclient.discovery import build
from httplib2 import Http
import time
import datetime
import pytz
import asyncio

from pytz import timezone

from oauth2client.service_account import ServiceAccountCredentials
from discord.ext import commands, tasks

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

basedir = os.path.abspath(os.path.dirname(__file__))
data_json = basedir+'/client_secret.json'

creds = ServiceAccountCredentials.from_json_keyfile_name(data_json, scope)
gc = gspread.authorize(creds)

shite = gc.open('Tempo')
rostersheet = shite.worksheet('WoE Roster')
#crsheet = shite.worksheet('Change Requests')
fullofsheet = shite.worksheet('Full IGNs')

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

############################## DEBUGMODE ##############################
debugger = False

################ Cell placements ###########################
guild_range = "B3:E99"
roster_range = "G3:J50"
matk_range = "L3:M14"
p1role_range = "P3:P14"
atk_range = "L17:M28"
p2role_range = "P17:P28"
p3_range = "L32:M43"
p3role_range = "P32:P43"
fullidname_range = "B4:C100"

############### Roles #######################################
list_ab = ['ab', 'arch bishop', 'arch', 'bishop', 'priest', 'healer', 'buffer']
list_doram = ['cat', 'doram']
list_gene = ['gene', 'genetic']
list_gx = ['gx', 'guillotine cross', 'glt. cross']
list_kage = ['kagerou', 'kage']
list_mech = ['mech', 'mechanic', 'mado']
list_mins = ['mins', 'minstrel' ]
list_obo = ['obo', 'oboro', 'ninja']
list_ranger = ['ranger', 'range']
list_rebel = ['rebel', 'reb', 'rebellion']
list_rg = ['rg', 'royal guard', 'devo',]
list_rk = ['rk', 'rune knight', 'db']
list_sc = ['sc', 'shadow chaser']
list_se = ['se', 'star emperor', 'hater']
list_sorc = ['sorc', 'sorcerer']
list_sr = ['sr', 'soul reaper', 'linker']
list_sura = ['sura', 'shura', 'asura', 'ashura']
list_wand = ['wanderer', 'wand', 'wandie', 'wandy']
list_wl = ['wl', 'warlock', 'tetra', 'crimson rock', 'cr']
      
  
############# Responses #####################################
answeryes = ['y', 'yes', 'ya', 'yup', 'ye', 'in', 'g']
answerno = ['n', 'no', 'nah', 'na', 'nope', 'nuh']

######################### CELERY RESPONSES ####################

answerzeny = ['zeny', 'zen', 'money', 'moneh', 'moolah']
answer10 = ['10', 'ten', 'plus ten', 'plusten', '10food', '+10', 'plustens', 'plus tens', '+10s']
answer20 = ['20', 'twenty', 'plus twenty', 'plustwenty', '20food', '+20', 'plustwentys', 'plus twentys', '+20s']
answernone = ['none', 'nada', 'nah', 'nothing', 'waive', 'waived']
answerevery = ['everything', 'all']
answerstr10 = ['+10 str', '+10str']
answeragi10 = ['+10 agi', '+10agi']
answervit10 = ['+10 vit', '+10vit']
answerint10 = ['+10 int', '+10int']
answerdex10 = ['+10 dex', '+10dex']
answerluk10 = ['+10 luk', '+10luk']
answerstr20 = ['+20 str', '+20str']
answeragi20 = ['+20 agi', '+20agi']
answervit20 = ['+20 vit', '+20vit']
answerint20 = ['+20 int', '+20int']
answerdex20 = ['+20 dex', '+20dex']
answerluk20 = ['+20 luk', '+20luk']
answerwhites = ['whites', 'hp pots', 'siege whites', 'white', 'siege white']
answerblues = ['blues', 'sp pots', 'siege blues', 'blue', 'siege blue']


############################# FEEDBACKS #############################

feedback_attplz = '```Please use /att y/n, y/n to register your attendance.```'
feedback_celeryplz = '```Please use /celery to list your salary preferences.```'
feedback_properplz = 'Please send a proper syntax: '
feedback_debug = '`[DEBUGINFO] `'




def next_available_row(sheet, column, lastrow):
    cols = sheet.range(3, column, lastrow, column)
    try:
        return max([cell.row for cell in cols if cell.value]) + 1
    except Exception as e:
        print(f'Handled exception: {e}, returning 3rd row')
        return 3


def next_available_row_p1(sheet, column):
    cols = sheet.range(3, column, 14, column)
    return max([cell.row for cell in cols if cell.value]) + 1


def next_available_row_p2(sheet, column):
    cols = sheet.range(17, column, 28, column)
    return max([cell.row for cell in cols if cell.value]) + 1


def next_available_row_p3(sheet, column):
    cols = sheet.range(32, column, 43, column)
    return max([cell.row for cell in cols if cell.value]) + 1


def sortsheet(sheet):
    issuccessful = True
    try:
        if sheet == rostersheet: 
            rostersheet.sort((4, 'asc'), (3, 'asc'), range=guild_range)
        elif sheet == celesheet:
            celesheet.sort((3, 'asc'), range = "B3:T99")
        elif sheet == silk2:
            silk2.sort((4, 'des'), (3, 'asc'), (2, 'asc'), range="B4:E51")
        elif sheet == silk4:
            silk4.sort((4, 'des'), (3, 'asc'), (2, 'asc'), range="B4:E51")
        #elif sheet == crsheet:
            #crsheet.sort((5, 'asc'), range="A3:G100")
        elif sheet == fullofsheet:
            fullofsheet.sort((5, 'asc'), (4, 'asc'), range="B4:H100")
        else:
            issuccessful = False
    except Exception as e:
        print(f'Exception caught at sortsheet: {e}')
        issuccessful = False
    return issuccessful


async def autosort(ctx, sheet):
    try: # Auto-sort
        issuccessful = sortsheet(sheet)
        if debugger: await ctx.send(f'{feedback_debug} Sorting {sheet.title} issuccessful={issuccessful}')
    except Exception as e:
        print(e)
        await ctx.send(f'{feedback_debug} Error on sorting {sheet.title}: `{e}`')
    return


def get_jobname(input):
    if input.lower() in list_ab:
        jobname = 'AB'
    elif input.lower() in list_doram:
        jobname = 'Doram'
    elif input.lower() in list_gene:
        jobname = 'Genetic'
    elif input.lower() in list_gx:
        jobname = 'GX'
    elif input.lower() in list_kage:
        jobname = 'Kagerou'
    elif input.lower() in list_mech:
        jobname = 'Mado'
    elif input.lower() in list_mins:
        jobname = 'Minstrel'
    elif input.lower() in list_obo:
        jobname = 'Oboro'
    elif input.lower() in list_ranger:
        jobname = 'Ranger'
    elif input.lower() in list_rebel:
        jobname = 'Rebel'
    elif input.lower() in list_rg:
        jobname = 'RG'
    elif input.lower() in list_rk:
        jobname = 'RK'
    elif input.lower() in list_sc:
        jobname = 'SC'
    elif input.lower() in list_se:
        jobname = 'Star Emperor'
    elif input.lower() in list_sorc:
        jobname = 'Sorc'
    elif input.lower() in list_sr:
        jobname = 'Soul Reaper'
    elif input.lower() in list_sura:
        jobname = 'Sura'
    elif input.lower() in list_wand:
        jobname = 'Wandie'
    elif input.lower() in list_wl:
        jobname = 'WL'
    else:
        jobname = ''
    return jobname


def reminder():
    attlist = [item for item in rostersheet.col_values(7) if item and item != 'IGN' and item != 'Next WOE:']
    ignlist = [item for item in rostersheet.col_values(3) if item and item != 'IGN' and item != 'READ THE NOTES AT [README]']
    row = 3
    dsctag = []
    dscid = []
    
    for ign in ignlist:
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
        else:
            gottem = 0
        row += 1
        
    return dsctag


class Clears(commands.Cog):
    def __init__(self, client):
        self.client = client

    # get debugmode
    def get_debugmode(self):
        return debugger

    @commands.command()
    async def tryto(self, ctx):
        cols = rostersheet.range(3, 2, 99, 2)
        #print(f'{cols}')
        try:
            print(f'{max(cell.row for cell in cols if cell.value)}')
        except Exception as e:
            print(f'found exception {e}')
        
        

    @commands.command()
    async def remind(self, ctx):
        ignlist = [item for item in rostersheet.col_values(3) if item and item != 'IGN' and item != 'READ THE NOTES AT [README]']
        global debugger
        channel = ctx.message.channel
        commander_name = ctx.author.name
        commander = ctx.author
        
        if channel.id in botinit_id:
            msg = await ctx.send(f'`Parsing the list. Please refrain from entering other commands.`')
            
            remindlist = reminder()
            remindlist.sort()
            if debugger: await ctx.send(f'{feedback_debug} Parsing... {remindlist}')
            
            try:
                embeded = discord.Embed(title = "Reminder List", description = "A list of people who really should /att y/n, y/n immediately", color = 0x00FF00)
            except Exception as e:
                print(f'discord embed reminder returned {e}')
                if debugger: await ctx.send(f'{feedback_debug} Error: `{e}`')
                return
            x = 0
            remlist = ''

            for x in range(len(remindlist)):
                remlist += remindlist[x] + '\n'
            try:
                embeded.add_field(name="Discord Tag", value=f'{remlist}', inline=True)
            except Exception as e:
                print(f'add field reminder returned {e}')
                if debugger: await ctx.send(f'{feedback_debug} Error: `{e}`')
                return
            
            try:
                await ctx.send(embed=embeded)
            except Exception as e:
                print(f'send embed remind returned {e}')
                if debugger: await ctx.send(f'{feedback_debug} Error: `{e}`')
            
            await ctx.send(f'Currently there are `{len(remindlist)}` who have not registered their attendance. {round((len(remindlist)/len(ignlist))*100, 2)}% of our guild have not registered.')
            
            await msg.delete()
        else:
            await ctx.send(f'Wrong channel! Please use #bot.')

    # toggle debugmode
    @commands.command()
    async def debugmode(self, ctx):
        global debugger
        channel = ctx.message.channel
        commander_name = ctx.author.name
        commander = ctx.author
        if channel.id in botinit_id:
            if commander.id in authorized_id:
                try:
                    debugger = not debugger
                except Exception as e:
                    await ctx.send(e)
                await ctx.send(f'`Debugmode = {debugger}`')
            else:
                await ctx.send(f'*Nice try pleb.*')
        else:
            await ctx.send(f'Wrong channel! Please use #bot.')

    # update discord member IDs
    @commands.command()
    async def refreshid(self, ctx):
        guild = ctx.guild
        channel = ctx.message.channel
        commander_name = ctx.author.name
        commander = ctx.author
        
        if channel.id in botinit_id:
            if commander.id in authorized_id:
                try:
                    msgprogress = await ctx.send('Refreshing Discord IDs for all members in Adonis Roster...')
                    cell_list = fullofsheet.range("C4:C100")
                    next_row = 4

                    for member in guild.members:
                        #await ctx.send(f'{member}')
                        if member.bot:
                            continue
                        else:
                            fullofsheet.update_cell(next_row, 2, str(member.id))
                            fullofsheet.update_cell(next_row, 3, str(member.name))
                            if debugger: await ctx.send(f'{feedback_debug} Updating {cell.value} ID at [{next_row}, 2] to {member.id}')
                                #break
                    next_row += 1
                    await msgprogress.edit(content="Refreshing Discord IDs for all members in Adonis Roster... Completed.")
                except Exception as e:
                    await msgprogress.edit(content="Refreshing Discord IDs for all members in Adonis Roster... Failed.")
                    await ctx.send(e)
            else:
                await ctx.send(f'*Nice try pleb.*')
        else:
            await ctx.send(f'Wrong channel! Please use #bot.')

    @commands.command()
    async def clearguild(self, ctx):
        channel = ctx.message.channel
        commander_name = ctx.author.name
        commander = ctx.author
        if channel.id in botinit_id:
            if commander.id in authorized_id:
                cell_list = rostersheet.range(guild_range)
                for cell in cell_list:
                    cell.value = ""
                rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                await ctx.send(f'{commander_name} has cleared the guild list.')
            else:
                await ctx.send(f'This command is unavailable for you!')
        else:
            await ctx.send(f'Wrong channel! Please use #bot.')
        # sh.values_clear("Sheet1!B3:E50")

    @commands.command()
    async def clearroster(self, ctx):
        channel = ctx.message.channel
        commander_name = ctx.author.name
        commander = ctx.author
        if channel.id in botinit_id:
            if commander.id in authorized_id:
                cell_list = rostersheet.range(roster_range)

                for cell in cell_list:
                    cell.value = ""

                rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                await ctx.send(f'{commander_name} has cleared the WoE Roster.')
            else:
                await ctx.send(f'This command is unavailable for you!')
        else:
            await ctx.send(f'Wrong channel! Please use #bot.')

    @commands.command()
    async def clearparty(self, ctx):
        channel = ctx.message.channel
        commander_name = ctx.author.name
        commander = ctx.author
        if channel.id in botinit_id:
            if commander.id in authorized_id:
                cell_list = rostersheet.range(matk_range)

                for cell in cell_list:
                    cell.value = ""

                #rostersheet.update_cells(cell_list)

                rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                cell_list = rostersheet.range(p1role_range)

                for cell in cell_list:
                    cell.value = ""

                rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                cell_list = rostersheet.range(atk_range)

                for cell in cell_list:
                    cell.value = ""

                rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                cell_list = rostersheet.range(p2role_range)

                for cell in cell_list:
                    cell.value = ""

                rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                cell_list = rostersheet.range(p3_range)

                for cell in cell_list:
                    cell.value = ""

                rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                cell_list = rostersheet.range(p3role_range)

                for cell in cell_list:
                    cell.value = ""

                rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')

                await ctx.send(f'{commander_name} has cleared the Party List.')
            else:
                await ctx.send(f'This command is unavailable for you!')
        else:
            await ctx.send(f'Wrong channel! Please use #bot.')

    @commands.command()
    async def enlist(self, ctx, *, arguments):
        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        if channel.id in botinit_id:
            arglist = [x.strip() for x in arguments.split(',')]
            no_of_args = len(arglist)
            if no_of_args < 2:
                await ctx.send(f'{ctx.message.author.mention} {feedback_properplz}`/enlist IGN, role, (optional comment)`')
                return
            else:
                darole = get_jobname(arglist[1])
                if darole == '':
                    await ctx.send(f'''Here are the allowed classes: 
```
For Doram: {list_doram}
For Genetic: {list_gene}
For Mechanic: {list_mech}
For Minstrel: {list_mins}
For Ranger: {list_ranger}
For Sorcerer: {list_sorc}
For Oboro: {list_obo}
For Rebellion: {list_rebel}
For Wanderer: {list_wand}
```
                                    ''')
                    return
                change = 0
                next_row = 3
                cell_list = rostersheet.range("B3:B99")
                for cell in cell_list:
                    if cell.value == commander_name:
                        change = 1
                        ign = rostersheet.cell(next_row, 3)
                        break
                    next_row += 1
                if change == 0:
                    #await ctx.send(f'testinggg')
                    #try:
                    next_row = next_available_row(rostersheet, 2, 99)
                    #except Exception as e:
                    #    await ctx.send(f'Failed: {e}')
                #await ctx.send(f'test')
                count = 0

                cell_list = rostersheet.range(next_row, 2, next_row, 5)
                for cell in cell_list:
                    if count == 0:
                        cell.value = commander_name
                    elif count == 1:
                        cell.value = arglist[0]
                    elif count == 2:
                        cell.value = darole
                    elif count == 3:
                        if no_of_args > 2:
                            cell.value = arglist[2]
                            optionalcomment = f', and Comment: {arglist[2]}'
                        else:
                            cell.value = ""
                            optionalcomment = ""
                    count += 1
                    
                #await ctx.send(f'test2')
                rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                await ctx.send(f'```{ctx.author.name} has enlisted {darole} with IGN: {arglist[0]}{optionalcomment}.```')
                if change == 1:
                    finding_column2 = celesheet.range("C3:C50".format(celesheet.row_count))
                    finding_columnsilk2 = silk2.range("B4:B51".format(silk2.row_count))
                    finding_columnsilk4 = silk4.range("B4:B51".format(silk4.row_count))
                    foundign2 = [found for found in finding_column2 if found.value == ign.value]
                    foundignsilk2 = [found for found in finding_columnsilk2 if found.value == ign.value]
                    foundignsilk4 = [found for found in finding_columnsilk4 if found.value == ign.value]

                    if foundignsilk2:
                        cell_list = silk2.range(foundignsilk2[0].row, 2, foundignsilk2[0].row, 4)
                        for cell in cell_list:
                            cell.value = ""
                        silk2.update_cells(cell_list, value_input_option='USER_ENTERED')
                        change = 0
                    if foundignsilk4:
                        cell_list = silk4.range(foundignsilk4[0].row, 2, foundignsilk4[0].row, 4)
                        for cell in cell_list:
                            cell.value = ""
                        silk4.update_cells(cell_list, value_input_option='USER_ENTERED')
                        change = 0
                    # Notify only once for any missing attendance
                    if foundignsilk2 or foundignsilk4:
                        await ctx.send(
                            f'{ctx.message.author.mention}``` I found another character of yours that answered for attendance already, I have cleared that. Please use /att y/n, y/n again in order to register your attendance.```')

                    if foundign2:
                        cell_list = celesheet.range(foundign2[0].row, 2, foundign2[0].row, 20)
                        for cell in cell_list:
                            cell.value = ""
                        celesheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                        #await ctx.send(f'``` I found another character of yours that answered celery preferences already, I have cleared that. Please use /celery again in order to list your preferred salary.```')
                        change = 0
                    else:
                        if not foundignsilk4 or not foundignsilk2:
                            await ctx.send(f'{feedback_attplz}')
                        #if not foundign2:
                            #await ctx.send(f'{feedback_celeryplz}')
                            
                        change = 0
                else:
                    await ctx.send(f'{ctx.message.author.mention} {feedback_attplz}')
                    #await ctx.send(f'{feedback_celeryplz}')
            await autosort(ctx, rostersheet)
            await autosort(ctx, celesheet)
            await autosort(ctx, silk2)
            await autosort(ctx, silk4)
        else:
            await ctx.send("Wrong channel! Please use #bot.")

    @commands.command()
    async def att(self, ctx, *, arguments):
        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        
        if not channel.id in botinit_id:
            await ctx.send("Wrong channel! Please use #bot.")
            return
        
        arglist = [x.strip() for x in arguments.split(',')]
        no_of_args = len(arglist)
        if (no_of_args != 1
                or not (arglist[0].lower() in answeryes or arglist[0].lower() in answerno)
            ):
            await ctx.send(f'{feedback_properplz} `/att y/n` *E.g. `/att y` to confirm attend both Silk 2')
            return
        
        next_row = 3
        found = 0
        cell_list = rostersheet.range("B3:B99")
        for cell in cell_list:
            if cell.value == commander_name:
                found = 1
                break
            next_row += 1
        if found == 0:
            await ctx.send(f'{ctx.message.author.mention} You have not yet enlisted your character. Please enlist via: `/enlist IGN, class, (optional comment)`')
            return
        
        ign = rostersheet.cell(next_row, 3)
        role = rostersheet.cell(next_row, 4)

        finding_column2 = silk2.range("B3:B50".format(rostersheet.row_count))
        finding_column4 = silk4.range("B3:B50".format(rostersheet.row_count))

        foundign2 = [found for found in finding_column2 if found.value == ign.value]
        foundign4 = [found for found in finding_column4 if found.value == ign.value]

        try:
            if foundign2:
                change_row = foundign2[0].row
            else:
                try:
                    change_row = next_available_row(silk2, 2, 51)
                except ValueError as e:
                    change_row = 4
            if debugger: await ctx.send(f'{feedback_debug} SILK 2 change_row=`{change_row}`')
            cell_list = silk2.range(change_row, 2, change_row, 4)
            count = 0
            # await ctx.send('test2')
            for cell in cell_list:
                # await ctx.send(f'test3 {ign.value} {role.value} {count}')
                if count == 0:
                    # await ctx.send(f'test4 {ign.value} {role.value} {count}')
                    cell.value = ign.value
                elif count == 1:
                    # await ctx.send(f'test5 {ign.value} {role.value} {count}')
                    cell.value = role.value
                elif count == 2:
                    # await ctx.send(f'test6 {ign.value} {role.value} {count}')
                    if arglist[0].lower() in answeryes:
                        cell.value = 'Yes'
                        yes = 1
                    else:
                        cell.value = 'No'
                    re_answer = cell.value
                count += 1
        except Exception as e:
            await ctx.send(f'Error on SILK 2: `{e}`')
            return
        
        # Ignore silk 2 entry if entered between post-silk 2 and pre-silk 4 time
        isskip = False
        try:
            my_time = pytz.timezone('Asia/Kuala_Lumpur')
            my_time_unformatted = datetime.datetime.now(my_time)
            my_dow = my_time_unformatted.strftime('%A')
            my_timeonly = my_time_unformatted.time()
            woeendtime = datetime.time(0, 0) # both silk 2 and 4 end on 00:00:00
            if debugger: await ctx.send(f'{feedback_debug} dayofweek=`{my_dow}` timeonly=`{my_timeonly}` woeendtime=`{woeendtime}`')
            if my_dow == 'Sunday' and my_timeonly >= woeendtime:
                isskip = True
        except Exception as e:
            await ctx.send(f'Time check error: `{e}`')
        
        if isskip:
            await ctx.send(f'```Ignoring {ctx.author.name}\'s answer {re_answer} for SILK 2 as the WoE for this week has already passed.```')
        else:
            silk2.update_cells(cell_list, value_input_option='USER_ENTERED')
            await ctx.send(f'```{ctx.author.name} said {re_answer} for SILK 2 with IGN: {ign.value}, Class: {role.value}.```')
        yes = 0
        try:
            if foundign4:
                change_row = foundign4[0].row
            else:
                try:
                    change_row = next_available_row(silk4, 2, 51)
                except ValueError as e:
                    change_row = 4
                cell_list = silk4.range(change_row, 2, change_row, 4)
            if debugger: await ctx.send(f'{feedback_debug} SILK 4 change_row=`{change_row}`')
            cell_list = silk4.range(change_row, 2, change_row, 4)
            count = 0
            # await ctx.send('test2')
            for cell in cell_list:
                # await ctx.send(f'test3 {ign.value} {role.value} {count}')
                if count == 0:
                    # await ctx.send(f'test4 {ign.value} {role.value} {count}')
                    cell.value = ign.value
                elif count == 1:
                    # await ctx.send(f'test5 {ign.value} {role.value} {count}')
                    cell.value = role.value
                elif count == 2:
                    # await ctx.send(f'test6 {ign.value} {role.value} {count}')
                    if arglist[1].lower() in answeryes:
                        cell.value = 'Yes'
                        yes = 1
                    else:
                        cell.value = 'No'
                    re_answer = cell.value
                count += 1
            silk4.update_cells(cell_list, value_input_option='USER_ENTERED')
            await ctx.send(f'```{ctx.author.name} said {re_answer} for SILK 4 with IGN: {ign.value}, Class: {role.value}.```')
        except Exception as e:
            await ctx.send(f'Error on SILK 4: `{e}`')
            return
        await autosort(ctx, silk2)
        await autosort(ctx, silk4)

    @commands.command()
    async def list(self, ctx):
        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        if channel.id in botinit_id:
            try:
                row_n = next_available_row(rostersheet, 7, 99)
            except ValueError:
                row_n = 3
            try:
                row_c = next_available_row(rostersheet, 8, 99)
            except ValueError:
                row_c = 3
            try:
                row_a = next_available_row(rostersheet, 9, 99)
            except ValueError:
                row_a = 3
            msg = await ctx.send(f'`Please wait... I am parsing a list of our WOE Roster. Refrain from entering any other commands.`')
            while row_n != row_c or row_n != row_a:
                row_n = next_available_row(rostersheet, 7, 99)
                row_c = next_available_row(rostersheet, 8, 99)
                row_a = next_available_row(rostersheet, 9, 99)

                if row_n < row_c:
                    if row_n < row_a:
                        cell_list = rostersheet.range(row_n, 7, row_n, 9)
                        for cell in cell_list:
                            cell.value = ""
                        rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                    else:
                        cell_list = rostersheet.range(row_a, 7, row_a, 9)
                        for cell in cell_list:
                            cell.value = ""
                        rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                elif row_c < row_a:
                    cell_list = rostersheet.range(row_c, 7, row_c, 9)
                    for cell in cell_list:
                        cell.value = ""
                    rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')
                else:
                    cell_list = rostersheet.range(row_a, 7, row_a, 9)
                    for cell in cell_list:
                        cell.value = ""
                    rostersheet.update_cells(cell_list, value_input_option='USER_ENTERED')
            try:
                namae = [item for item in rostersheet.col_values(7) if item and item != 'IGN' and item != 'Next WOE:']
            except Exception as e:
                print(f'namae returned {e}')
            try:
                kurasu = [item for item in rostersheet.col_values(8) if item and item != 'Class' and item != 'Silk 2' and item != 'Silk 4']
            except Exception as e:
                print(f'kurasu returned {e}')
            try:
                stat = [item for item in rostersheet.col_values(9) if item and item != 'Att.']
            except Exception as e:
                print(f'stat returned {e}')
            #komento = [item for item in rostersheet.col_values(10) if item and item != 'Comments']
            x = 0
            a = 0
            yuppie = 0
            noppie = 0
            for a in stat:
                if a == 'Yes':
                    yuppie += 1
                else:
                    noppie += 1

            if yuppie == 0 and noppie == 0:
                await ctx.send(f'`Attendance not found. `\n{feedback_attplz}')
                await msg.delete()
                return

            try:
                embeded = discord.Embed(title = "Current WOE Roster", description = "A list of our Current WOE Roster", color = 0x00FF00)
            except Exception as e:
                print(f'discord embed returned {e}')
                return
            x = 0
            fullname = ''
            fullclass = ''
            fullstat = ''

            for x in range(len(namae)):
                fullname += namae[x] + '\n'
                fullclass += kurasu[x] + '\n'
                fullstat += stat[x] + '\n'
            try:
                embeded.add_field(name="IGN", value=f'{fullname}', inline=True)
            except Exception as e:
                print(f'add field returned {e}')
                return
            embeded.add_field(name="Class", value=f'{fullclass}', inline=True)
            try:
                embeded.add_field(name="Status", value=f'{fullstat}', inline=True)
            except Exception as e:
                print(f'add field returned {e}')
                return


            try:
                await ctx.send(embed=embeded)
            except Exception as e:
                print(f'send embed returned {e}')
            await ctx.send(f'Total no. of Yes answers: {yuppie}')
            await ctx.send(f'Total no. of No answers: {noppie}')
            await msg.delete()
        else:
            await ctx.send("Wrong channel! Please use #bot.")


    @commands.command()
    async def help(self, ctx):

        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        if channel.id in botinit_id:
            await ctx.send("""__**BOT COMMANDS**__
PLEASE MIND THE COMMA, IT ENSURES THAT I SEE EVERY ARGUMENT:

**/enlist** `IGN`, `class`, *`optional comment`*
> enlists your Discord ID, IGN, Class, and optional comment in the GSheets
> e.g. `/enlist Ayaneru, Sura`
**/att** `y/n`, `y/n`
> registers your attendance (either yes or no) in the GSheets, for silk 2 and 4 respectively.
> e.g. `/att y, n` *(to attend silk 2, skip silk 4)*
**/list**
> parses a list of the current attendance list
**/listpt**
> parses a list of the current party list divided into ATK, MATK, and SECOND GUILD
**/remind**
> lists down members who have yet to register their attendance.
""")
            if commander.id in authorized_id:
                msghelpadmin = '''
**/debugmode**
> For development use. Toggles debugging mode: some features will result in extra feedbacks with `[DEBUGINFO]`
> Some features will behave differently during debugmode.
**/clearguild**
> clears guild list
**/clearroster**
> clears attendance list
**/clearparty**
> clears party list
**/forcetimedevent `name`, `time`**
> **name** = timed event name - one of the following: archive, remind1, remind2, reset
> **time** = time to schedule, in the format of hh:mm:ss:Day. Case sensitive!
**/refreshid**
> updates Discord ID of all members in the list'''
                await ctx.send(f'Hi boss! Here are the **admin-only commands**:{msghelpadmin}')
        else:
            await ctx.send("Wrong channel! Please use #bot.")

    @commands.command()
    async def listpt(self, ctx):
        channel = ctx.message.channel
        commander = ctx.author
        commander_name = commander.name
        if channel.id in botinit_id:
            msg = await ctx.send(
                f'`Please wait... I am parsing a list of our Party List. Refrain from entering any other commands.`')
            cell_list = rostersheet.range("M4:M15")
            get_MATK = [""]
            for cell in cell_list:
                get_MATK.append(cell.value)
            cell_list = rostersheet.range("M19:M30")
            get_ATK = [""]
            for cell in cell_list:
                get_ATK.append(cell.value)
            cell_list = rostersheet.range("M34:M45")
            get_third = [""]
            for cell in cell_list:
                get_third.append(cell.value)

            MATK_names = [item for item in get_MATK if item]
            ATK_names = [item for item in get_ATK if item]
            THIRD_names = [item for item in get_third if item]

            try:
                embeded = discord.Embed(title="Current Party List", description="A list of our Current Party List",
                                        color=0x00FF00)
            except Exception as e:
                print(f'discord embed returned {e}')
                return
            x = 0
            ATKpt = ''
            MATKpt = ''
            THIRDpt = ''
            for x in range(len(MATK_names)):
                MATKpt += MATK_names[x] + '\n'
            x = 0
            for x in range(len(ATK_names)):
                ATKpt += ATK_names[x] + '\n'
            x = 0
            for x in range(len(THIRD_names)):
                THIRDpt += THIRD_names[x] + '\n'
            try:
                embeded.add_field(name="ATK Party", value=f'{ATKpt}', inline=True)
            except Exception as e:
                print(f'add field returned {e}')
                return
            embeded.add_field(name="MATK Party", value=f'{MATKpt}', inline=True)
            try:
                embeded.add_field(name="SECOND GUILD Party", value=f'{THIRDpt}', inline=True)
            except Exception as e:
                print(f'add field returned {e}')
                return
            try:
                await ctx.send(embed=embeded)
            except Exception as e:
                print(f'send embed returned {e}')
            await msg.delete()
            # return
        else:
            await ctx.send("Wrong channel! Please use #bot.")

def setup(client):
    client.add_cog(Clears(client))
