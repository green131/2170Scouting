import time
import datetime
from scouting import clock, date

print "Initializing"
original_doc = open('C:\Users\Daniel\Desktop\original.txt', 'rU')
new_doc = open('C:/Users/Daniel/Desktop/new.txt', 'a')
new_doc.write('START:')
x = 0
for line in original_doc:
    if '#FRC' in str(line):
        line = line.split('TY')[1]
        new_doc.write(str(date()) + '-' + str(clock(0)) + ' ' + 'TY' + line)
        x += 1
        print 'Wrote Line ' + str(x)
    else:
        pass
original_doc.close()
new_doc.close()