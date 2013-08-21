def mean(data):
    i = 0
    sum = 0
    while i < len(data):
        sum = data[i] + sum
        i = i + 1
    else:
        mean = sum/len(data)
    return mean

def median(data):
    data.sort()
    median = data[len(data)/2]
    return median

def mode(data):
    i = 0
    n = 0
    c = 0
    v = 0
    data.sort()
    while i < len(data):
        if int(data[i]) == int(n):
            c += 1
        elif int(data[i]) != int(n) and int(data[i - 1]) != int(n):
            v += 1
            if c < v:
                mode = data[i]
                print mode
            else:
                break
        else:
            v = 0
        i += 1
    return mode

def standard_deviation(data):
    from math import sqrt
    i = 0
    n = 0
    while i < len(data):
        dev = ((data[i] - mean(data)) ** 2)
        n = n + dev
        i = i + 1
    st_dev = round(sqrt(n / (len(data) - 1)), 3)
    return st_dev

#-------------------Main Script-------------------------
if __name__ == '__main__':
    #Python - Number List Functions
    #Main data set:
    data = [1, 2, 5, 3, 3, 7, 2, 2, 2, 5]
    #Loop:
    l = 0
    while l < 1:
    #Input Info
        print "The data set is: " + str(data)
        print "What do you want to do?"
        print "1 : Find the mean of the data"
        print "2 : Find the median of the data"
        print "3 : Find the mode of the data"
        print "4 : Find the standard deviation of the data"
        #User Input
        user_choice = str(raw_input("Function Choice: "))
        if user_choice == "1":
            print "The mean of the data is " + str(mean(data)) + "."
        elif user_choice == "2":
            print "The median of the data is: " + str(median(data)) + "."
        elif user_choice == "3":
            print "The mode of the data is " + str(mode(data)) + "."
        elif user_choice == "4":
            print "The standard deviation of the data is " + str(standard_deviation(data)) + "."
        else:
            print "Please input a correct choice."
        exit = str(raw_input("Would you like to do another function? (Y/N) "))
        if exit.lower() == "y":
            l = 1
        elif exit.lower() == "n":
            break
        else:
            print "Please input a correct choice."

