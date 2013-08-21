#-------------------------------------------------------------------------------
# Name:        Team 2170 Email Reader
# Purpose:     Email Reading System for Scouting Program
#
# Author:      Daniel A. Green
#
# Created:     12/08/2011
# Copyright:   (c) Daniel 2011
# Licence:     CCR License (Not for distrubution, sharing, or editing in any way unless authorized by the Author)
#-------------------------------------------------------------------------------

from scouting import *
from opr_ranking import *

#Option Menu
repeat_toggle = True            #Whether the program repeats.
repeat_time = 15                #Time (in seconds) before repeat.
message_counter = True          #Whether the program display on host machine the number of new msgs.
admin_calls = True              #Whether an admin can perform special actions.
admin_list = ['dan.singkiwi@gmail.com','8609703398@mms.att.net']
opr_ranking_active = True       #Enables OPR Gathering / Ranking

#Send Message Format:
#send_email(server, '2170scouting@gmail.com (ME)', 'RECIPIENT', 'SUBJECT', 'BODY')

#Location of Local Excel WorkBook
file_name = 'C:/Users/Daniel/Desktop/My Documents/Robotics/2012_Global_Scouting_System/scouting_data.xlsx'

#PROGRAM START
repeat = True
connection_error_wait = 0
opr_timer_end = datetime.datetime.now() + datetime.timedelta(minutes=60)
while repeat == True:
    #toggle repeat on / off
    repeat = repeat_toggle
    print '(' + str(clock(0)) + ') Email Grabber Initiated (Repeat == ' + str(repeat) + ')'
    pop_conn, server = connect('2170scouting@gmail.com', 'bDBnMG0wdGkwbg==')
    #if no connection: pass
    if pop_conn == False:
        pass
    else:
        connection_error_wait = 0
        messages = get_messages(pop_conn)
        if message_counter == True:
            #Optional New Message Counter:
            if len(messages) == 0:
                print '>>> You have no new messages.'
            else:
                if len(messages) > 1: c = 'messages.'
                else: c = 'message.'
                print '>>> You have', len(messages), 'new', c
                #End Optional Message Counter
        if len(messages) == 0:
            pass
        else:
            sheets, workBook = excel_open(file_name)
            #if workbook error: pass
            if sheets == False:
                pass
            else:
                sent_to_sender = []
                format_sender = []
                pending = []
                #Format individual message
                for message in messages:
                    error = None
                    #Check if whitelisted
                    sheet = excel_sheet('whitelist', sheets)
                    whitelisted, Row, sender = check_sender(message, sheet)
                    #if not whitelisted, check if pending
                    if whitelisted == False:
                        sheet = excel_sheet('pending', sheets)
                        currently_pending, Row = find_term(sheet, 'C', sender)
                    #if data availible, pull it
                    if whitelisted == True or currently_pending == True:
                        name = str(sheet.Cells(Row, 'D').Value)
                        team = int(float(sheet.Cells(Row, 'E').Value))
                    #format subject
                    subject = format_subject(message)
                    #format body
##                    try:
                    content, msg, match = format_body(format_msg(message, subject))
##                    except:
##                        content = []
##                        msg = []
##                        match = None
                    #ADMIN CALLS
                    #Option to enable / disable admin calls
                    if str(sender) in admin_list:
                        if admin_calls == True:
                            #Admin Shuts down program
                            if re.search('(.)*'+'end program'+'(.)*', str(msg)) != None:
                                #Marks Admin Call
                                content.append('Admin Call')
                                error = 'Admin Terminated Program'
                                repeat = False
                                send_email(server, '2170scouting@gmail.com', sender, 'Program Terminating',
                                'Program has reciceved a termination request. Shutting down.')
                            #Admin Checks Pending
                            elif re.search('(.)*'+'check pending'+'(.)*', str(msg)) != None:
                                content.append('Admin Call')
                                pending_users = 'Pending Entries:'
                                sheet = excel_sheet('pending', sheets)
                                x = 0
                                Row = 2
                                entries = {}
                                confirm = 0
                                empty_counter = 0
                                full_counter = 0
                                while x < 1:
                                    if empty_counter > 50:
                                        x = 1
                                    elif sheet.Cells(Row,'A').Value == None:
                                        Row += 1
                                        empty_counter += 1
                                    else:
                                        empty_counter = 0
                                        full_counter += 1
                                        entries[full_counter] = str(sheet.Cells(Row, 'C').Value)
                                        line = (str(str(sheet.Cells(Row, 'B').Value).split(' ')[0]) + ' | ' +
                                                    str(sheet.Cells(Row, 'C').Value) + ' | ' +
                                                    str(sheet.Cells(Row, 'D').Value) + ' | ' +
                                                    str(float(int(sheet.Cells(Row, 'E').Value))))
                                        pending_users += '\n' + str(str(full_counter) + ': ' + str(line))
                                        Row += 1
                                print '(' + str(clock(0)) + ') Check Pending Called'
                                send_email(server, '2170scouting@gmail.com', str(sender), None, pending_users)
                            #Admin Confirms Pending
                            elif re.search('(.)*'+'confirm pending'+'(.)*', str(msg)) != None:
                                    #Marks Admin Call
                                    content.append('Admin Call')
                                    msg = str(msg).split('confirm pending')
                                    msg = ' '.join(msg[1].split())
                                    msg = re.match('[\s,\d]+', msg)
                                    if msg == None:
                                        print '\nError: Confirm Pending Format Error\n'
                                        send_email(server, '2170scouting@gmail.com', str(sender), None,
                                        ('Error: Confirm Pending Incorrectly Formatted.\n'
                                        'Format should be "Confirm Pending 1 2 3 4". '
                                        'Type "Check Pending" to check pending.'))
                                    else:
                                        msg = msg.group(0)
                                        print '(' + str(clock(0)) + ') Confirmation Program Initiated'
                                        if len(msg) > 0:
                                            msg = msg.split(' ')
                                            print '>>> Transferring ' + str(' '.join(msg)) + ' to Whitelist...'
                                            success = ['Users Confirmed:\n']
                                            failed = ['\nFailed:\n']
                                            #iterate through pending -> find entry [,/]
                                            for Row in msg:
                                                if re.match('(\d+)', str(Row)) != None:
                                                    Row = int(Row) + 1
                                                if re.match('(\S+)', str(Row)) == None:
                                                    pass
                                                else:
                                                    sheet = excel_sheet('pending', sheets)
                                                    #identify and extract data
                                                    email = sheet.Cells(Row, 'C').Value
                                                    name = sheet.Cells(Row, 'D').Value
                                                    team = sheet.Cells(Row, 'E').Value
                                                    if email == None or name == None or team == None or str(Row) <= '1':
                                                        failed.append(Row)
                                                        pass
                                                    else:
                                                        #Copy data to whitelist
                                                        sheet = excel_sheet('whitelist', sheets)
                                                        whitelist_add(sheet, email, name, team)
                                                        print '>>> Entry ' + str(name) + ' (' + str(email) + ') successfuly added to whitelist.'
                                                        #Erase row in pending
                                                        sheet = excel_sheet('pending', sheets)
                                                        sheet.Cells(Row, 'A').Value = None
                                                        sheet.Cells(Row, 'B').Value = None
                                                        sheet.Cells(Row, 'C').Value = None
                                                        sheet.Cells(Row, 'D').Value = None
                                                        sheet.Cells(Row, 'E').Value = None
                                                        success.append(Row)
                                                        send_email(server, '2170scouting@gmail.com', email, None,
                                                            'Registration Complete!\n'
                                                            'You have been registered with 2170scouting and can now submit scouting data. '
                                                            'New to the service? Visit our site for more information: http://goo.gl/FhxtQ')
                                        if len(success) == 1:
                                            success = 'No Users Were Confirmed'
                                        else:
                                            x = 0
                                            for part in success:
                                                if x != 0:
                                                    success[x] = str(int(part) - 1)
                                                x += 1
                                            success = str(' '.join(success))
                                        if len(failed) == 1:
                                            failed = ' '
                                        else:
                                            x = 0
                                            for part in failed:
                                                if x != 0:
                                                    failed[x] = str(int(part) - 1)
                                                x += 1
                                            failed = str(' '.join(failed))
                                        send_email(server, '2170scouting@gmail.com', str(sender), None, success + failed)
                    #Prevent an Admin Call from Passing other Functions
                    if 'Admin Call' not in content:
                        #WHITELISTED
                        if whitelisted == True:
                            #No match data
                            if len(content) == 0 or match == False:
                                #Help Call
                                if re.search('(.)*'+'help'+'(.)*', str(msg)) != None:
                                    content.append('Help Requested')
                                    error = 'Help Requested'
                                    #Make sure more than 1 message isn't sent
                                    if sender not in sent_to_sender:
                                        sent_to_sender.append(sender)
                                        send_email(server, '2170scouting@gmail.com', sender, None, ('Help Informaton\n'
                                            '2170scouting allows teams to easily submit scouting data.\n'
                                            'To report a problem or request further assistance, visit http://goo.gl/INXtD'))
                                #Check a teams OPR
                                elif re.search('(.)*'+'opr'+'(\s+)'+'(\d+)'+'(\s+)'+'(.)*', str(msg)) != None:
                                    content.append('Requested Team OPR')
                                    error = 'Team OPR Requested'
                                    #Make sure more than 1 message isn't sent
                                    if sender not in sent_to_sender:
                                        opr_team = None
                                        opr_team_ranking = None
                                        opr_team_overall_ranking = None
                                        sent_to_sender.append(sender)
                                        #Read requested team number
                                        msg = msg.split('opr')[1:]
                                        for part in msg:
                                            try:
                                                opr_team = re.match('(\s+)'+'(\d)+', part)
                                                opr_team = int(opr_team.group(0))
                                            except:
                                                pass
                                            #found opr team number
                                            if opr_team != None:
                                                break
                                        #no opr team number found
                                        if opr_team == None:
                                            error = 'Format Error: Could not identify Team to lookup OPR'
                                            send_email(server, '2170scouting@gmail.com', sender, None, ('Sorry!\n'
                                                'We were unable to read your opr request. Please \n'
                                                'For more info visit our help website at http://goo.gl/INXtD'))
                                        else:
                                            #Lookup of team opr ranking based on team number
                                            sheet = excel_sheet('opr', sheets)
                                            found, Row = find_term(sheet, 'A', opr_team)
                                            if found == True:
                                                opr_team_ranking = sheet.Cells(Row, 'B').Value
                                                opr_team_overall_ranking = int(Row) - 1
                                                #Find total number of teams
                                                found, Row = find_term(sheet, 'B', None)
                                                opr_total_scored_teams = int(Row) - 1
                                            #found opr team ranking
                                            if opr_team_ranking != None and opr_team_overall_ranking != None:
                                                send_email(server, '2170scouting@gmail.com', sender, None, ('OPR Rankings\n'
                                                    'Team ' + str(opr_team) + ' has an OPR of\n' + str(opr_team_ranking) +
                                                    ' and is ranked ' + str(opr_team_overall_ranking) + ' out of ' + str(opr_total_scored_teams)
                                                    + ' teams.'))
                                            else:
                                                error = 'Information Missing: Could not find OPR ranking for team ' + str(opr_team)
                                                send_email(server, '2170scouting@gmail.com', sender, None, ('Sorry!\n'
                                                    'We are unable to find a current OPR ranking of team ' + str(opr_team) + '.\n'
                                                    'For more info visit our help website at http://goo.gl/INXtD'))
                                #Read comments
                                elif re.search('(.)*'+'read comments'+'(\s+)'+'(\d+)'+'(.)*', str(msg)) != None:
                                    content.append('Reading Comments')
                                    error = 'Reading Comments'
                                    #Make sure more than 1 message isn't sent
                                    if sender not in sent_to_sender:
                                        sent_to_sender.append(sender)
                                        #Read team number
                                        part = msg.split('read comments')[1]
                                        team = None
                                        try:
                                            team = re.match('(\s+)'+'(\d+)', part)
                                            team = float(team.group(0))
                                        except:
                                            team = None
                                        if team != None:
                                            sheet = excel_sheet('comments', sheets)
                                            found, Row = find_term(sheet, 'A', float(int(team)))
                                            if found == True:
                                                total_comments = ''
                                                comment = sheet.Cells(Row, 'B').Value
                                                comment = comment.split(' | ')
                                                for part in comment:
                                                    total_comments += str(part) + '\n'
                                                send_email(server, '2170scouting@gmail.com', sender, None, (
                                                    'Comments for team ' + str(int(team)) + ':\n' + total_comments))
                                            else:
                                                send_email(server, '2170scouting@gmail.com', sender, None, ('Sorry!\n'
                                                    'We were unable to find any comments for team ' + str(int(team)) + '.'))
                                        else:
                                            send_email(server, '2170scouting@gmail.com', sender, None, ('Sorry!\n'
                                                'We were unable to read your comment request. Please '
                                                'check your formatting and try again.\n'
                                                'For more info visit our help website at http://goo.gl/INXtD'))
                                #Add a comment
                                elif re.search('(.)*'+'comment'+'(\s+)'+'(\d+)'+'(\s+)'+'(.)*'+'..'+'(.)*', str(msg)) != None:
                                    content.append('New Comment')
                                    error = 'New Comment'
                                    #Make sure more than 1 message isn't sent
                                    if sender not in sent_to_sender:
                                        sent_to_sender.append(sender)
                                        #Read team number
                                        team = None
                                        msg = str(msg).split('comment')[1:]
                                        for part in msg:
                                            part = part.split('..')
                                            for section in part:
                                                if re.search('(\s+)'+'(\d+)'+'(\s+)'+'(.)*', section) != None:
                                                    try:
                                                        team = re.match('(\s+)'+'(\d+)', section)
                                                        team = float(int(team.group(0)))
                                                        comment = section[:]
                                                    except:
                                                        team = None
                                        if team != None:
                                            comment = comment.split(str(int(team)))
                                            comment = comment[-1]
                                            if type(comment) == str:
                                                sheet = excel_sheet('comments', sheets)
                                                found, Row = find_term(sheet, 'A', float(int(team)))
                                                if found == True:
                                                    sheet.Cells(Row, 'B').Value += ' | ' + str(comment)
                                                else:
                                                    sheet.Cells(Row, 'A').Value = str(int(team))
                                                    sheet.Cells(Row, 'B').Value = str(comment)
                                                send_email(server, '2170scouting@gmail.com', sender, None, (
                                                    'Comment successfuly added for team ' + str(int(team)) + ' . To '
                                                    'retrive just send "read comments ##[team number]##".\n'
                                                    'For more info visit our help website at http://goo.gl/INXtD'))
                                            else:
                                                send_email(server, '2170scouting@gmail.com', sender, None, ('Sorry!\n'
                                                    'We were unable to read your comment. Please '
                                                    'check your formatting and try again.\n'
                                                    'For more info visit our help website at http://goo.gl/INXtD'))
                                        else:
                                            send_email(server, '2170scouting@gmail.com', sender, None, ('Sorry!\n'
                                                'We were unable to read your comment. Please '
                                                'check your formatting and try again.\n'
                                                'For more info visit our help website at http://goo.gl/INXtD'))
                                #Predict match outcome
                                elif re.search('(.)*'+'predict'+'(\s+)'+'(\d+)'+'(\s+)'+'(\d+)'+'(\s+)'+'(\d+)'+'(\s+)'+'vs'
                                    +'(\s+)'+'(\d+)'+'(\s+)'+'(\d+)'+'(\s+)'+'(\d+)'+'(\s+)'+'(.)*', str(msg)) != None:
                                    content.append('Requested Match Prediction')
                                    error = 'Match Prediction Requested'
                                    #Make sure more than 1 message isn't sent
                                    if sender not in sent_to_sender:
                                        teams = None
                                        sent_to_sender.append(sender)
                                        #Read team numbers
                                        message = message.split('predict')[1:]
                                        for part in message:
                                            try:
                                                teams = re.match('(\s+)'+'(\d+)'+'(\s+)'+'(\d+)'+'(\s+)'+'(\d+)'+'(\s+)'+'vs'
                                                    +'(\s+)'+'(\d+)'+'(\s+)'+'(\d+)'+'(\s+)'+'(\d+)'+'(\s+)', part)
                                                teams = teams.group(0)
                                                #Pull out individual team numbers
                                                individual_teams = []
                                                teams = teams.split('vs')
                                                teams[0] = teams[0].split(' ')
                                                teams[1] = teams[1].split(' ')
                                                for part in teams[0]:
                                                    if re.search('(\d+)', part) != None:
                                                        match = re.match('(\d+)', part)
                                                        individual_teams.append(match.group(0))
                                                for part in teams[1]:
                                                    if re.search('(\d+)', part) != None:
                                                        match = re.match('(\d+)', part)
                                                        individual_teams.append(match.group(0))
                                            except:
                                                pass
                                            #found opr team number
                                            if len(individual_teams) == 6:
                                                break

                                        #no opr team number found
                                        if len(individual_teams) != 6:
                                            error = 'Format Error: Could not predict match'
                                            send_email(server, '2170scouting@gmail.com', sender, None, ('Sorry!\n'
                                                'We were unable to read your match prediction request. Please '
                                                'check your formatting and try again.\n'
                                                'For more info visit our help website at http://goo.gl/INXtD'))
                                        else:
                                            #Lookup of each team opr ranking based on team number
                                            team_OPRs = {}
                                            sheet = excel_sheet('opr', sheets)
                                            for team in individual_teams:
                                                found, Row = find_term(sheet, 'A', team)
                                                if found == True:
                                                    opr_team_ranking = sheet.Cells(Row, 'B').Value
                                                    team_OPRs[team] = opr_team_ranking
                                                else:
                                                    break
                                        if len(team_OPRs) < 6:
                                            error = 'Information Missing: Could not find OPR ranking for team ' + str(team)
                                            send_email(server, '2170scouting@gmail.com', sender, None, ('Sorry!\n'
                                                    'We are unable to find a current OPR ranking of team ' + str(team) + '.\n'
                                                    'For more info visit our help website at http://goo.gl/INXtD'))
                                        else:
                                            team1 = 0
                                            team2 = 0
                                            x = 0
                                            for team in individual_teams:
                                                if x < 3:
                                                    team1 += team_OPRs[team]
                                                else:
                                                    team2 += team_OPRs[team]
                                                x += 1
                                            if team1 > team2:
                                                difference = team1 / (team1 + team2)
                                                winners = individual_teams[0:3]
                                            else:
                                                difference = team2 / (team1 + team2)
                                                winners = individual_teams[3:5]
                                            winners = str(winners[0] + ', ' + winners[1] + ', and ' + winners[2])
                                            send_email(server, '2170scouting@gmail.com', sender, None, ('Match Prediction\n'
                                                'Teams ' + str(winners) + ' will win by a ' + str(round(difference, 3) * 100) + '%' +
                                                ' difference.'))

                                #Format Error Call
                                else:
                                    content.append(str(msg))
                                    error = 'Format Error'
                                    #Make sure more than 1 message isn't sent
                                    if sender not in format_sender:
                                        format_sender.append(sender)
                                        send_email(server, '2170scouting@gmail.com', sender, 'Unable to Process',
                                            'WHOOPS! Site is unable to process your data submission.\n'
                                            'Make sure your data is in the correct format:\n'
                                            'match ## team #### score ##\n'
                                            'Reply help for more info.')
                            #Record Message Data and Resulting Actions
                            counter = 0
                            for item in content:
                                msg_list = []
                                msg_list.append(str(item))
                                if error == None:
                                    sheet = excel_sheet('data', sheets)
                                    #Count number of data submissions
                                    counter += 1
                                else:
                                    sheet = excel_sheet('error', sheets)
                                output_print(sender, name, team, subject, error, str(msg_list))
                                output_write(sheet, sender, name, team, subject, error, str(msg_list))
                            #Compose set submission email
                            if counter >= 1:
                                if counter > 1: set = 'sets'
                                elif counter == 1: set = 'set'
                                send_email(server, '2170scouting@gmail.com', sender, None, ('Data Submitted\n'
                                    'You have sucessfully submitted ' + str(counter) + ' ' + set + ' of data.'))
                        #NOT WHITELISTED, Found registration data
                        #Find if name + team info in msg
                        elif currently_pending == False and re.search('(.)*'+'(name)'+'(\s+)'+'(\w+)'+'(\s+)'+'(\w+)*'+'(.)*'+'(\d+)', str(msg)) != None:
                            names = str(msg).split('name')
                            for item in names:
                                #find location of info
                                if re.match('(\s)'+'(\w+)'+'(\s)'+'(\w+)', str(item)):
                                    first = item.split()[0]
                                    first = first.capitalize()
                                    last = item.split()[1]
                                    last = last.capitalize()
                                    name = str(first) + ' ' + str(last)
                                    for iterable in item.split():
                                        if re.search('(\d+)', str(iterable)) != None:
                                            team = re.match('(\d+)', str(iterable))
                                            team = team.group(0)
                                            break
                                    break
                            Row = whitelist_add(sheet, sender, name, team)
                            print '(' + str(clock(0)) + ') User ' + str(sender) + ' added to pending list as ' + str(name) + ' (' + str(team) + ')'
                            send_email(server, '2170scouting@gmail.com', sender, None, ('Registration Pending\n'
                                'We will contact you when your registration is complete so you can start scouting! '
                                'Need help or have a question? Visit our help site for more information: http://goo.gl/INXtD'))
                            #Add Pending Name to Pending List
                            pending.append('< ' + str(Row) + ' | ' + str(sender) + ' | ' + str(name) + ' | ' + str(team) + ' >')
                        #NOT WHITELISTED, No registration data (Differs if pending)
                        else:
                            if currently_pending == True:
                                sheet = excel_sheet('error', sheets)
                                error = 'Sender Currently in Pending List'
                                #format msg to one line
                                msg_list = []
                                msg_list.append(msg)
                                output_write(sheet, sender, name, team, subject, error, str(msg_list))
                                print '(' + str(clock(0)) + ') User ' + str(sender) + ' already in pending list as ' + str(name) + ' (' + str(team) + ')'
                                send_email(server, '2170scouting@gmail.com', sender, None, ('Sorry, your registration is currently pending.\n'
                                    'We will contact you when your registration is complete so you can start scouting! '
                                    'Need help or have a question? Visit our help site for more information: http://goo.gl/INXtD'))
                            else:
                                print '(' + str(clock(0)) + ') Rejected: Unknown sender ' + '(' + sender + ') must provide info'
                                send_email(server, '2170scouting@gmail.com', sender, 'Unregistered User', ('This site requires all users to be registered.\n'
                                    'Please register by replying to this message with your details in the following format:\n\'name &&&& &&&& team ####\'\n'
                                    'Reply help for more info.'))
                    #Send Admin Pending Registrations
                    pending_users = ''
                    if len(pending) != 0:
                        for names in pending:
                            pending_users = pending_users + str(names + '\n')
                        send_email(server, '2170scouting@gmail.com', 'dan.singkiwi@gmail.com', 'Pending Registrations', ('Users Pending:\n' + str(pending_users) + '\n'))
                excel_close(workBook)
        connection_quit(pop_conn, server)
    #OPR
    if opr_ranking_active == True:
        print '(' + str(clock(0)) + ') Checking for new match data . . .'
        filehandle, new_match_data = opr_get_data()
        #Check if internet connection / new match data:
        #(set to 'or' to rebuild OPR ranking, keep at 'and' otherwise)
        if filehandle != False and opr_timer_end < datetime.datetime.now():
            print '(' + str(clock(0)) + ') Initializing OPR Calculator'
            OPR = opr_calculator()
            if OPR != None:
                sheets, workBook = excel_open(file_name)
                sheet = excel_sheet('opr', sheets)
                print '(' + str(clock(0)) + ') Writing New OPR Values (this may take a while) . . .'
                output_write_clear(sheet, ['A', 'B'])
                for team in OPR:
                    if int(team[0]) > 0:
                        print '>>> Writing team ' + str(team)
                        if abs(int(team[1])) < 1000:
                            output_write_simple(sheet, int(team[0]), significant_figures(team[1], 3))
                print '(' + str(clock(0)) + ') OPR Recalculated'
                excel_close(workBook)
            opr_timer_end = datetime.datetime.now() + datetime.timedelta(minutes=60)
    #record start time
    start_time = datetime.datetime.now()
    #repeat
    if repeat == True:
        if pop_conn == False and connection_error_wait < 1200:
            connection_error_wait += 120
        print '(' + str(clock(0)) + ') Waiting for', repeat_time + connection_error_wait, 'seconds. You may now exit the program if necessary.'
        end_time = start_time + datetime.timedelta(seconds=(repeat_time + connection_error_wait))
        hold = True
        import time
        while hold == True:
            if datetime.datetime.now() > end_time:
                hold = False
            else:
                time.sleep((repeat_time + connection_error_wait)/4)
        print ' '
print '(' + str(clock(0)) + ') Shutting Down Email Grabber'