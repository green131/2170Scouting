#-------------------------------------------------------------------------------
# Name:        Team 2170 Scouting
# Purpose:     Def Module for Scouting Program
#
# Author:      Daniel
#
# Created:     30/12/2011
# Copyright:   (c) Daniel 2011
# Licence:     CCR License (Not for distrubution, sharing, or editing in any way unless authorized by the Author)
#-------------------------------------------------------------------------------
#!/usr/bin/env python

import poplib
import email
import smtplib
import re
import os
import sys
import time
import threading
import base64
import datetime
from win32com.client import Dispatch
from email import parser
from email.mime.text import MIMEText
from smtplib import SMTP_SSL as SMTP

def check_sender(message, sheet):
    whitelisted = False
    sender = str(message['from']).split('<')
    if len(sender) > 1:
        sender = sender[1]
        sender = sender.split('>')
        sender = str(sender[0])
    else:
        sender = str(sender[0])
    x = 0
    Row = 2
    while x < 1:
        if sheet.Cells(Row,'C').Value == sender:
            whitelisted = True
            break
        elif sheet.Cells(Row,'C').Value == None:
            whitelisted = False
            break
        else:
            Row += 1
    return whitelisted, Row, sender

def clock(seconds_forward):
    time = datetime.datetime.now()
    if len(str(time.hour)) == 1:
        hour = str(0) + str(time.hour)
    else:
        hour = time.hour
    if len(str(time.minute)) == 1:
        minute = str(0) + str(time.minute)
    else:
        minute = time.minute
    total_seconds = int(time.second) + int(seconds_forward)
    if len(str(total_seconds)) == 1:
        second = str(0) + str(time.second)
    else:
        second = total_seconds
    clock = str(hour) + ':' + str(minute) + '.' + str(second)
    return clock

def cls():
    os.system(['clear','cls'][os.name == 'nt'])

def confirm_sender(locations, sheet_list):

    return success

def connect(username, password):
    print '(' + str(clock(0)) + ') Logging In . . .'
    try:
        #pop3
        pop_conn = poplib.POP3_SSL('pop.gmail.com')
        pop_conn.user(username)
        pop_conn.pass_(base64.b64decode(password))
        #smtp
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(username, base64.b64decode(password))
        print '(' + str(clock(0)) + ') Login Successful'
        return pop_conn, server
    except:
        print '(' + str(clock(0)) + ') ERROR: Could not get new emails [Unable to connect to google server]'
        return False, False

def connection_quit(pop_conn, server):
    try:
        pop_conn.quit()
        server.quit()
        print '(' + str(clock(0)) + ') Quit Connections'
    except:
        print '(' + str(clock(0)) + 'ERROR: Could not disconnect!'

def date():
    time = datetime.datetime.now()
    if len(str(time.month)) == 1:
        month = str(0) + str(time.month)
    else:
        month = time.month
    if len(str(time.day)) == 1:
        day = str(0) + str(time.day)
    else:
        day = time.day
    date = str(time.month) + '/' + str(time.day) + '/' + str(time.year)
    return date

def excel_close(workBook):
    try:
        workBook.Saved = 0
        workBook.Save()
        workBook.Close(SaveChanges=True)
    except:
        print '(' + str(clock(0)) + ') ERROR: Could not close Workbook'

def excel_open(file_name):
    try:
        excel = Dispatch('Excel.Application')
        excel.Visible = False #If we want to see it change
        data_sheet = excel.Workbooks.Open(file_name)
        #Activate Excel Parts
        workBook = excel.ActiveWorkbook
        activeSheet = excel.ActiveSheet
        sheets = workBook.Sheets
        return sheets, workBook
    except:
        print '(' + str(clock(0)) + ') ERROR: Failed to initalize workbook. Retrying...'
        try:
            excel_close(workBook)
        except:
            pass
        try:
            excel = Dispatch('Excel.Application')
            excel.Visible = False #If we want to see it change
            data_sheet = excel.Workbooks.Open(file_name)
            #Activate Excel Parts
            workBook = excel.ActiveWorkbook
            activeSheet = excel.ActiveSheet
            sheets = workBook.Sheets
            return sheets, workBook
        except:
            print '(' + str(clock(0)) + ') ERROR: Failed to initialize workbook on second try.'
            return False, False

def excel_sheet(choice, sheets):
    #Indentify and Activate Exel Sheet
    sheet = sheets(choice)
    sheet.Activate()
    return sheet

def find_term(sheet, column, term):
    found = False
    x = 0
    Row = 2
    while x < 1:
        if sheet.Cells(Row, column).Value == term:
            found = True
            break
        elif sheet.Cells(Row, column).Value == None:
            found = False
            break
        else:
            Row = Row + 1
    return found, Row

def format_match(message):
    match = False
    try:
        match = re.match('(.)*' + '(match)' + '(\d+)', str(message))
    except:
        pass
    return match

def format_body(message):
    match = False
    content = []
    x = 0
    for part in message:
        if x < 1 or len(part) <= 1:
            x += 1
        else:
            if len(part) > 1:
                y = 0
                for part2 in part:
                    if len(part2) < 1:
                        y += 1
                        pass
                    elif y == 0:
                        try:
                            match = re.match('(.)*' + '(\d+)', part2)
                            match = match.group(0)
                            y += 1
                        except AttributeError:
                            match = False
                    else:
                        try:
                            #identify sets of data ([team#] [hybrid total]
                            #    [tele-total] [balancing total]
                            section = re.match(('(.)*' + '(\d+)'+ '(\s+)' + '(ht)'+ '(\s+)' + '(\d+)'
                                + '(\s+)' + '(tt)' + '(\s+)' + '(\d+)' + '(\s+)' + '(bt)' + '(\s+)' + '(\d+)'), part2)
                            section = section.group(0)
                            section = ' '.join(section.split())
                            content.append(str('match ' + match + 'team ' + section))
                            y += 1
                        except AttributeError:
                            pass
    content = list(set(content))
    if len(message) < 2 and type(message) == str:
        message = ' '.join(message)
    elif len(message) < 2 and type(message) == list:
        message = message[0]
    if len(message) == 1 and type(message) == list:
        message = message[0]
    return content, message, match

def format_msg(message, subject):
    msg = 'Null'
    msg = str(subject) + ' ' + str(message)
    msg = msg.lower()
    msg = msg.split('match')
    # msg[x,[x,y]]
    x = 0
    for part in msg:
        if 'team' in part:
            msg[x] = part.split('team')
        x += 1
    return msg

def format_subject(message):
    try:
        subject = message['subject']
        if len(subject) == 0: subject = 'None'
    except:
        subject = 'None'
    return subject

def get_messages(pop_conn):
    #Get messages from server:
    messages = [pop_conn.retr(i) for i in range(1, len(pop_conn.list()[1]) + 1)]
    #Concat message pieces:
    messages = ['\n'.join(mssg[1]) for mssg in messages]
    #Parse message into an email object:
    messages = [parser.Parser().parsestr(mssg) for mssg in messages]
    return messages

def get_plays(team, alliance_pairings):
    count = 0
    for alliance in alliance_pairings:
        if team in alliance:
            count += 1
    return count

def output_print(sender, name, team, subject, error, content):
    print ' '
    print '>>> Date   :', date()
    print '>>> Time   :', clock(0)
    print '>>> From   :', sender
    print '>>> Name   :', name
    print '>>> Team   :', team
    print '>>> Subject:', subject
    if error != None:
        print '>>> Error  :', error, '(Check Error Log For Details)'
    else:
        print '>>> Content:', content
    print ' '

def output_write(sheet, sender, name, team, subject, error, content):
    Row = 2
    x = 0
    while x < 1:
        if sheet.Cells(Row,'A').Value == None:
            sheet.Cells(Row, 'A').Value = clock(0)
            sheet.Cells(Row, 'B').Value = date()
            sheet.Cells(Row, 'C').Value = sender
            sheet.Cells(Row, 'D').Value = name
            sheet.Cells(Row, 'E').Value = team
            sheet.Cells(Row, 'F').Value = subject
            sheet.Cells(Row, 'G').Value = content
            if error != None:
                sheet.Cells(Row, 'H').Value = error
            break
        else:
            Row += 1

def output_write_clear(sheet, columns):
    for column in columns:
        Row = 2
        x = 0
        while x < 1:
            if sheet.Cells(Row, column).Value == None:
                break
            else:
                sheet.Cells(Row, column).Value = ''
                Row += 1

def output_write_simple(sheet, team, value):
    Row = 2
    x = 0
    while x < 1:
        if sheet.Cells(Row,'A').Value == None:
            sheet.Cells(Row, 'A').Value = team
            sheet.Cells(Row, 'B').Value = value
            break
        else:
            Row += 1

def purge_team(searched_team, alliance_pairings, final_scores, teams):
    for alliance in alliance_pairings:
        if searched_team in alliance:
            del final_scores[alliance_pairings.index(alliance)]
            del alliance_pairings[alliance_pairings.index(alliance)]
            for team in alliance:
                if get_plays(team, alliance_pairings) < 4 and team != searched_team:
                    purge_team(team, alliance_pairings, final_scores, teams)
    try:
        del teams[teams.index(searched_team)]
    except:
        pass
    return teams, alliance_pairings, final_scores

def send_email(server, me, receiver, subject, body):
    text_subtype = 'plain'
    msg = MIMEText(body, text_subtype)
    msg['From'] = me
    msg['Subject'] = subject
    server.sendmail(me, receiver, msg.as_string())

def significant_figures(x, n):
             # Use %e format to get the n most significant digits, as a string.
             format = "%." + str(n-1) + "e"
             as_string = format % x
             return as_string

def whitelist_add(sheet, email, name, team):
    Row = 2
    x = 0
    while x < 1:
        if sheet.Cells(Row,'A').Value == None:
            sheet.Cells(Row, 'A').Value = clock(0)
            sheet.Cells(Row, 'B').Value = date()
            sheet.Cells(Row, 'C').Value = email
            sheet.Cells(Row, 'D').Value = name
            sheet.Cells(Row, 'E').Value = team
            break
        else:
            Row += 1
    return Row