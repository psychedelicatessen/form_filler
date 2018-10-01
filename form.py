################################################
# Author: https://github.com/psychedelicatessen#
################################################

import openpyxl

# open the template and select the payment sheet for editing
wb = openpyxl.load_workbook('type1.xlsx')

sheet = wb['Payment Form']

# define the commandline prompt
prompt = "> "

# define a class constructor for assignments
class Assignment:

    def __init__(self, a_name, a_time="", a_date=[], hours=None, pay_rate=None, activity=""):

        self.name = a_name

        self.time = a_time

        self.date = a_date

        self.hours = hours

        self.rate = pay_rate

        self.activdsg = activity

# define seperate functions for assignment variables with potentially varying data
    def addTime(self):
        
        added_time = input(prompt)

        self.time = added_time

    def addDate(self):

        user_input = input(prompt)

        added_date = list(user_input.split(','))

        self.date = added_date

    def addHours(self):

        added_hours = input(prompt)

        self.hours = added_hours
    def addRate(self):

        pay_rate = input(prompt)

        self.rate = pay_rate

    def activ(self):

        activity = input(prompt)

        self.activdsg = activity

# define assignment constructor function
def create_assignment(assName):

    assignment = Assignment(assName)

    print("Please input the assignment's time period.")
    assignment.addTime()

    print("Please input the assignment's date. (a list is okay here, seperated by commas. i.e. 09-11,09-12,09-13, etc.)")
    assignment.addDate()

    print("Please input total hours worked (per day).")
    assignment.addHours()
    
    print("Please specify the assignment's payrate.")
    assignment.addRate()

    print("Please specify the assignment's activity designation.")
    assignment.activ()

    return(assignment)

# define a class constructor function for the transportation sheet repurpose the class's properties for the sheet
def create_transport(transName):

    transport = Assignment(transName)

    print("Please input the assignment's location.")
    # use the time property to store the location
    transport.addTime()

    print("Please input the Transportation type.")
    # use the activity property to store the transport type
    transport.activ()

    print("Please input the date of travel. (a list is okay here, seperated by commas. i.e. 05-01,05-02,05-03, etc.)")
    transport.addDate()

    print("Please add the number of visits.")
    # use the hours property to store the # of visits
    transport.addHours()

    print("Please add the cost of travel.")
    transport.addRate()

    return(transport)

# define a reusable function to fill the rows of the spreadsheet
def fill_row():
    for i in fields:
        if i == 1:
            sheet.cell(row=position, column=i).value = assignment.name
        elif i == 2:
            sheet.cell(row=position, column=i).value = assignment.activdsg
        elif i == 3:
            sheet.cell(row=position, column=i).value = assignment.time
        elif i == 4:
            sheet.cell(row=position, column=i).value = assignment.hours
        elif i == 5:
            sheet.cell(row=position, column=i).value = assignment.date[count]
        elif i == 6:
            sheet.cell(row=position, column=i).value = assignment.hours
        else:
            sheet.cell(row=position, column=i).value = assignment.rate

# refactored function to fill the rows of the transport sheet
def fill_row_transport():
    for i in transfields:
        if i == 2:
            sheet.cell(row=position, column=i).value = transport.name
        elif i == 3:
            sheet.cell(row=position, column=i).value = transport.time
        elif i == 4:
            sheet.cell(row=position, column=i).value = transport.activdsg
        elif i == 5:
            sheet.cell(row=position, column=i).value = transport.date[count]
        elif i == 7:
            sheet.cell(row=position, column=i).value = transport.hours
        elif i == 8:
            sheet.cell(row=position, column=i).value = transport.rate


# ask for preliminary information
print("Please enter your name.")

name = input(prompt)

sheet['A5'] = 'Instructor: ' + name

print(f"\nHello, {name}. Please enter the Pay Period (MM-DD MM-DD i.e. 07-01 07-16)")

period = input(prompt)

sheet['D5'] = 'Period: ' + period

print(f"\n{period}? Okay. Now enter the submission date.")

subdate = input(prompt)

sheet['F5'] = 'Submitted: ' + subdate

##################################################
#move on to collecting the assignment information#
##################################################

print('\n----------')
print('Let\'s move on, shall we?')
print('----------')

print('Define a new assignment? (Y/N only)')

# store the user's answer in the answer variable
answer = input(prompt)

# set the row position with an integer
position = 8

# define the number of columns in a 'fields' array for easy iterating
fields = [1,2,3,4,5,6,7]

# count variable for multiple dates
count = 0

# continuous loop which populates the spreadsheet with the provided assignment information.
while True:

    if answer == 'Y':

        count = 0

        print('Input the name of the assignment')

        companyName = input(prompt)

        assignment = create_assignment(companyName)
        
        # after creation of the class object check to see if there are multiple dates in the class. If so, iterate through them and push them to the spreadsheet, else just fill out the row and iterate the position.

        if len(assignment.date) > 1:

            while count < len(assignment.date):
                
                fill_row()
                
                position += 1

                count += 1
        else:

            fill_row()

            position += 1

        print('\nCreate another assignment?')
        answer =input(prompt)

    elif answer == 'N':

        print('Okay, all done.')
        break 
    else:

        print('Y or N only please.')

        answer = input(prompt)

# switch to the transporation sheet for editing

print('Now let\'s fill out the transportation section.')

# ensure count is reset
count = 0

# reset position value
position = 8

# redefine fields for the stupidly laid out transporation sheet
transfields = [2,3,4,5,6,7,8]

sheet = wb['Transportation']


# Set the top fields for the transportation page

sheet['B5'] = 'Instructor: ' + name
sheet['D5'] = 'Period: ' + period
sheet['G5'] = 'Submitted: ' + subdate


print('Enter a new transportation transaction?')

answer = input(prompt)

while True:

    if answer == 'Y':

        count = 0

        print('Input the name of the assignment')

        transportName = input(prompt)

        transport = create_transport(transportName)
        
        # after creation of the class object check to see if there are multiple dates in the class. If so, iterate through them and push them to the spreadsheet, else just fill out the row and iterate the position.

        if len(transport.date) > 1:

            while count < len(transport.date):
                
                fill_row_transport()
                
                position += 1

                count += 1
        else:

            fill_row_transport()

            position += 1

        print('\nDid you buy a drink?')
        drink_answer = input(prompt)

        if drink_answer == 'Y':
            print('How much did you pay?')
            amount = input(prompt)
            
            sheet.cell(row=position, column=2).value = transport.name
            sheet.cell(row=position, column=3).value = 'Drink'
            sheet.cell(row=position, column=5).value = transport.date[count]
            sheet.cell(row=position, column=7).value = transport.hours
            sheet.cell(row=position, column=8).value = amount

            position += 1


        print('\nCreate another transaction?')
        answer = input(prompt)

    elif answer == 'N':

        print('Okay, all done.')
        break 
    else:

        print('Y or N only please.')

        answer = input(prompt)

# format a string for the filename
name = name.replace(" ", "")
period = period.replace(" ", "")

title = name.strip() + period.strip() + '.xlsx'

print('Now saving as: ' + title + ' Goodbye.')
wb.save(title)
