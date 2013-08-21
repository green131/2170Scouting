from scouting import *

#PROGRAM START
print '>>> (' + str(clock()) + ') Confirmation Program Initiated'
#Location of Local Excel WorkBook
file_name = 'C:\Users\Daniel\Desktop\My Documents\Robotics\Global Scouting System\scouting_data.xlsx'
sheets, workBook = excel_open(file_name)
if sheets == False:
    pass
else:
    print 'Pending Entries:'
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
            print str(full_counter) + ': ' + str(line)
            Row += 1
    if full_counter != 0:
        print ('Type Entry Numbers You Wish to Confirm:'
                        '(separated by a space)')
        print ' '
        confirm = raw_input()
        if len(confirm) > 0:
            confirm = confirm.split(' ')
            print 'Transferring to Whitelist...'
            #iterate through pending -> find entry [,/]
            for item in confirm:
                x = 0
                Row = 2
                try:
                    item = int(item.encode('utf8'))
                    email = entries[item]
                except KeyError:
                    print 'Sorry, entry ' + str(item) + ' could not be found.'
                    x = 1
                while x < 1:
                    if sheet.Cells(Row,'C').Value == str(email):
                        #identify and extract data
                        name = sheet.Cells(Row, 'D').Value
                        team = sheet.Cells(Row, 'E').Value
                        #Copy data to whitelist
                        sheet = excel_sheet('whitelist', sheets)
                        whitelist_add(sheet, email, name, team)
                        pop_conn, server = connect('2170scouting@gmail.com','bDBnMG0wdGkwbg==')
                        print 'Entry ' + str(item) + ' (' + email + ') successfuly added to whitelist.'
                        send_email(server, '2170scouting@gmail.com', email, None,
                                        'Registration Complete!\n'
                                        'You have been registered with 2170scouting and can now submit scouting data. '
                                        'New to the service? Visit our site for more information: http://goo.gl/FhxtQ')
                        connection_quit(pop_conn, server)
                        #Erase row in pending
                        sheet = excel_sheet('pending', sheets)
                        sheet.Cells(Row, 'A').Value = None
                        sheet.Cells(Row, 'B').Value = None
                        sheet.Cells(Row, 'C').Value = None
                        sheet.Cells(Row, 'D').Value = None
                        sheet.Cells(Row, 'E').Value = None
                        x = 1
                        break
                    else:
                        Row += 1
    else:
        print 'There are no pending entries.'
excel_close(workBook)
try:
    close = input('Press ENTER to Close')
except:
    pass