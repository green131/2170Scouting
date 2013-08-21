#location of frcfms RSS feed
#https://twitter.com/#!/frcfms

#Tags
#TY, MC, RF, BF, RE, BL, RA, BA, RB, BB, RFP, BFP, RHS, BHS, RTS, BTS, BP, RP, CP

#Key
#TY match Type (P=practice, Q=qualification, E=elimination)
#MC Match number
#RF Red Final score
#BF Blue Final score
#RA Red Alliance teams
#BA Blue Alliance teams
#RB Red Balance score
#BB Blue Balance score
#RFP Red Foul Points
#BFP Blue Foul Points
#RHS Red Hybrid Score
#BHS Blue Hybrid Score
#RTS Red Teleoperated Score
#BTS Blue Teleoperated Score
#CP Coopertition Points

#sample
#data = 'TY Q MC 11 RF 42 BF 30 RA 3276 1111 69 BA 1413 333 57 RB 0 BB 0 RFP 42 BFP 30 RHS 0 BHS 0 RTS 0 BTS 0 CP 0'

import re
import shutil
import os
import time
import datetime
import math
import urllib
import operator
import numpy as numpy
from scouting import *

def opr_get_data():
    new_match_data = False
    try:
        filehandle = urllib.urlopen('https://api.twitter.com/1/statuses/user_timeline.rss?screen_name=frcfms')
    except:
        filehandle = False
        print '(' + str(clock(0)) + ') ERROR: Could not download new match data [Unable to connect to frcfms twitter feed]'
    if filehandle != False:
        #Read latest recorded data
        most_recent_year = False
        most_recent_month = False
        most_recent_day = False
        most_recent_hour = False
        most_recent_minute = False
        most_recent_second = False
        doc = open('C:/Users/Daniel/Desktop/My Documents/Robotics/2012_Global_Scouting_System/match_data.txt', 'rU')
        count = 0
        for line in doc:
            count += 1
            if count > 1000:
                pass
            most_recent_signal = False
            date_read = (line.split(' ')[0]).split('-')[0]
            date_read_year = int(date_read.split('/')[2])
            date_read_month = int(date_read.split('/')[0])
            date_read_date = int(date_read.split('/')[1])
            clock_read = (line.split(' ')[0]).split('-')[1]
            clock_read_hour = int(clock_read.split(':')[0])
            clock_read_minute = int((clock_read.split(':')[1]).split('.')[0])
            clock_read_second = int((clock_read.split(':')[1]).split('.')[1])
            if most_recent_year < date_read_year:
                most_recent_signal = True
            elif most_recent_year == date_read_year:
                if most_recent_month < date_read_month:
                    most_recent_signal = True
                elif most_recent_month == date_read_month:
                    if most_recent_day < date_read_date:
                        most_recent_signal = True
                    elif most_recent_day == date_read_date:
                        if most_recent_hour < clock_read_hour:
                            most_recent_signal = True
                        elif most_recent_hour == clock_read_hour:
                            if most_recent_minute < clock_read_minute:
                                most_recent_signal = True
                            elif most_recent_minute == clock_read_minute:
                                if most_recent_second < clock_read_second:
                                    most_recent_signal = True
            if most_recent_signal == True:
                most_recent_year = date_read_year
                most_recent_month = date_read_month
                most_recent_day = date_read_date
                most_recent_hour = clock_read_hour
                most_recent_minute = clock_read_minute
                most_recent_second = clock_read_second
                #Record most recent match data
                most_recent_data = line.split(' ')[1:]
                most_recent_data[-1] = most_recent_data[-1].split('\n')[0]
                most_recent_data = ' '.join(most_recent_data)
        doc.close
        #Append new data
        doc = open('C:/Users/Daniel/Desktop/My Documents/Robotics/2012_Global_Scouting_System/match_data.txt', 'a')
        x = 0
        for line in filehandle.readlines():
            if '<title>frcfms:' in line:
                x += 1
                formatted_line_data = line.split(' ')
                del formatted_line_data[:2]
                formatted_line_data[-1] = formatted_line_data[-1].split('</title>')[0]
                del formatted_line_data[0:4]
                formatted_line_data = ' '.join(formatted_line_data)
                #Reading new match data, append
                if x == 1:
                    seconds = 14
                else:
                    seconds = 0
                if formatted_line_data != most_recent_data:
                    doc.write('\n' + str(date()) + '-' + str(clock(seconds)) + ' ' + formatted_line_data)
                    new_match_data += 1
                #Hit most recent data, so quit lookup of data
                if formatted_line_data == most_recent_data:
                    break
        doc.close
        if new_match_data == False:
            print '>>> No new match data.'
        else:
            plural = 's'
            if new_match_data == 1:
                plural = ''
            print '>>> ' + str(new_match_data) + ' new set' + plural + ' of match data.'
    return filehandle, new_match_data

def opr_calculator():
    print '(' + str(clock(0)) + ') Reading Data'
    #Read all data for processing
    error = None
    doc = open('C:/Users/Daniel/Desktop/My Documents/Robotics/2012_Global_Scouting_System/match_data.txt', 'rU')
    alliance_pairings = []
    teams = []
    final_scores = []
    for line in doc:
        #Split data into list type
        data = line.split(' ')
        #Test type of match, only use qualification matches
        if str(data[2]) == 'Q':
            #Read match data to seperate variable
            raw_data = []
            x = 9
            del data[0]
            del data[12]
            #Append all teams once to teams list
            while x < 15:
                #Pick teams not to include
                if int(data[x]) in [0]:
                    pass
                elif data[x] not in teams:
                    teams.append(data[x])
                x += 1
            #Append all final scores (in order) to matrix  [CHOOSE ONE]
            #OPR
            final_scores.append([int(data[5])])
            final_scores.append([int(data[7])])
            #CCP
    ##        final_scores.append([int(data[5]) - int(data[7])])
    ##        final_scores.append([int(data[7]) - int(data[5])])
            #Append team pairings
            alliance_pairings.append(data[9:12]) #red alliance teams
            alliance_pairings.append(data[12:15]) #blue alliance teams
    doc.close

    print '(' + str(clock(0)) + ') Formatting Data'
    for team in teams:
    	if purge_team(team, alliance_pairings, final_scores, teams) < 3:
    		teams, alliance_pairings, final_scores = purge_team(team, alliance_pairings, final_scores, teams)

    x = 0
    while x < len(alliance_pairings):
        x += 1
        stay = {}
        for alliance in alliance_pairings:
            stay[str(alliance)] = False
            for team in alliance:
                if get_plays(team, alliance_pairings) == 3:
                    stay[str(alliance)] = True
        for alliance in alliance_pairings:
            if stay[str(alliance)] == False:
                del final_scores[alliance_pairings.index(alliance)]
                del alliance_pairings[alliance_pairings.index(alliance)]
                break

    while len(final_scores) > len(teams):
        del final_scores[-1]

##    #Print length of lists
##    print '---length of lists---'
##    print '---(must be equal)---'
##    print 'teams: ' + str(len(teams))
##    print 'alliance_pairings: ' + str(len(alliance_pairings))
##    print 'final_scores: ' + str(len(final_scores))
    #Build blank matrix
    x = 0
    matrixMatches = []
    while x < len(teams):
        matrix_rows = []
        y = 0
        while y < len(teams):
            matrix_rows.append(0)
            y += 1
        matrixMatches.append(matrix_rows)
        x += 1
    #Fill in blank matrix
    for team in teams:
        for lists in alliance_pairings:
            if team in lists:
                for alliance_buddy in lists:
                    matrixMatches[teams.index(team)][teams.index(alliance_buddy)] += 1

    print '(' + str(clock(0)) + ') Solving OPR'
    #SOLVE for OPR
    OPR = False
    try:
        matrix_matchups = numpy.matrix(matrixMatches)                    # Creates a matrix.
        matrix_final_scores = numpy.matrix(final_scores)                 # Creates a matrix (like a column vector).
        OPR = numpy.linalg.lstsq(matrix_matchups, matrix_final_scores)   # Solves linalg matrix equation
    except:
        OPR = 'OPR Could Not be Calculated'
        error = 'Incompatible matrix sizes'
        print '(' + str(clock(0)) + ') ERROR: Could not calculate OPR [Incompatible matrix dimensions / Need more data]'
    OPR_dict = {}
    sorted_OPR_dict = None
    if type(OPR) != str:
        #Format Matrix to new list
        x = 0
        for part in OPR[0]:
            OPR_dict[teams[x]] = round((part.tolist())[0][0], 2)
            x += 1
        #Sort OPR dict by values
        sorted_OPR_dict = sorted(OPR_dict.iteritems(), key=operator.itemgetter(1), reverse=True)
    return sorted_OPR_dict

if __name__ == '__main__':
    opr_get_data()
##    from list_functions import mean
##    x = 0
##    while x < 1:
##        OPR_values = []
##        print 'Calculating:'
##        x += 1
##        opr_calc = opr_calculator()
##        print opr_calc
##        print '---OPR Calculated---'
##        if opr_calc != None:
##            for item in opr_calc:
##                print item
##            print '---individual values---'
##            for item in opr_calc:
##                OPR_values.append(int(item[1]))
##                if '341' in item:
##                    print item
##                if '3658' in item:
##                    print item
##            print 'Mean: ' + str(mean(OPR_values))


##    temp = input('Press ENTER to exit')

