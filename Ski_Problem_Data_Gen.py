import xlsxwriter
import random

#Timerline ski hours are TYPICALLY:
# 9 am --> 4 pm 
# 7 hours: Data every half hour = 14 datapoints

# Workbook() takes one, non-optional, argument
# which is the filename that we want to create.
workbook = xlsxwriter.Workbook('Ski_Problem.xlsx')


# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
# The worksheet name will be "Ski Data"
worksheet = workbook.add_worksheet("Ski Data")


#Data titles we want to write:

Data_titles = ["Time", "Entering Park" , "Exiting Park" , "Season Ticket Holders", "Day-Pass Holders"]
Time = ["9 am", "9:30 am", "10 am", "10:30am ", "11 am", "11:30 am", "12 am", "12:30 pm", "1 pm", "1:30 pm", "2 pm", "2:30 pm", "3 pm", "3:30 pm", "4 pm"]

column = 0
row = 0

#Iterate through titles in the data list and write them
for title in Data_titles:
    worksheet.write(row, column, title)
    column +=1

row = 1   
#Iterate through times in Time list and write them
for t in Time:
    worksheet.write(row, 0, t)
    row +=1

#====================================================
#Creates data for entering park total: 3-parts
#Creates data for Entering in morning
entering_in_Morning = [] 
for i in range (4) :
    entering_in_Morning.append (random.randint(125, 200))

#Creates data for Entering in afternoon
entering_in_Afternoon = []
for i in range (5) :
    entering_in_Afternoon.append (random.randint(100, 150))

#Creates data for Entering at end of day
entering_in_End_of_Day = []
for i in range (5):
    entering_in_End_of_Day.append (random.randint(0, 50))

#Writes the data for Entering
column = 1
row = 1
for rate in entering_in_Morning:
    worksheet.write(row, column, rate)
    row +=1

for rate in entering_in_Afternoon:
    worksheet.write(row, column, rate)
    row +=1

for rate in entering_in_End_of_Day:
    worksheet.write(row, column, rate)
    row +=1

worksheet.write('B16', 0)


#=========================================================

#Creates data for people leaving park: 3-parts

leaving_in_Morning = [] 
for i in range (5) :
    leaving_in_Morning.append (random.randint(0, 35))

#Creates data for leaving in afternoon
leaving_in_Afternoon = []
for i in range (5) :
    leaving_in_Afternoon.append (random.randint(75, 150))

#Creates data for leaving at end of day
leaving_in_End_of_Day = []
for i in range (4):
    leaving_in_End_of_Day.append (random.randint(150, 180))


#Writes the data for Leaving Park
column = 2
row = 1
for rate in leaving_in_Morning:
    worksheet.write(row, column, rate)
    row +=1

for rate in leaving_in_Afternoon:
    worksheet.write(row, column, rate)
    row +=1

for rate in leaving_in_End_of_Day:
    worksheet.write(row, column, rate)
    row +=1

worksheet.write('C16', 0)
#==========================================================

#Creates data for people Season tickets in park

SeasonTicketsInPark = []
for i in range (14):
    SeasonTicketsInPark.append (random.randint(50, 75))

#writes the data to the spreadsheet
row = 1
for rate in SeasonTicketsInPark:
    worksheet.write(row, 3, rate)
    row+=1 

worksheet.write('D16', 0)

#=========================================================
#Finds totals of entering and avg per half hour per day

total_morning = 0
for i in entering_in_Morning :
    total_morning += i

total_afternoon = 0
for i in entering_in_Afternoon:
    total_afternoon += i

total_End_of_Day = 0
for i in entering_in_End_of_Day:
    total_End_of_Day += i

total_entering_Throughout_day = total_morning + total_afternoon + total_End_of_Day
average_entering_Throughout_day = total_entering_Throughout_day / 14

worksheet.write('A18', 'Average:')
worksheet.write('B18', average_entering_Throughout_day)
#=========================================================
#Finds totals for leaving and avg per half hour per day

total_exiting_morning = 0
for i in leaving_in_Morning:
    total_exiting_morning += i

total_exiting_afternoon = 0
for i in leaving_in_Afternoon:
    total_exiting_afternoon += i

total_exiting_end_of_day = 0
for i in leaving_in_End_of_Day:
    total_exiting_end_of_day += i

total_leaving_throughout_day = total_exiting_morning + total_exiting_afternoon + total_exiting_end_of_day
average_leaving_Throughout_day = total_leaving_throughout_day / 14

worksheet.write('C18', average_leaving_Throughout_day)
#=========================================================
#Creates data for day pass holders and writes data to spreadsheet
#Also creates the total daypass variable
total_day_pass_holders = 0
day_pass_holders = 0
row = 1
#column = 0


for x, y in zip(entering_in_Morning, SeasonTicketsInPark):
    day_pass_holders = x - y
    total_day_pass_holders += day_pass_holders
    if day_pass_holders < 0:
        worksheet.write(row, 4, 0)
        row+= 1
    else:
        worksheet.write(row, 4, day_pass_holders)
        row+= 1

for x, y in zip(entering_in_Afternoon, SeasonTicketsInPark):
    day_pass_holders = x - y
    total_day_pass_holders += day_pass_holders
    if day_pass_holders < 0:
        worksheet.write(row, 4, 0)
        row+= 1
    else:
        worksheet.write(row, 4, day_pass_holders)
        row+= 1

for x, y in zip(entering_in_End_of_Day, SeasonTicketsInPark):
    day_pass_holders = x - y
    if day_pass_holders < 0:
        worksheet.write(row, 4, 0)
        row+= 1
    else:
        worksheet.write(row, 4, day_pass_holders)
        row+= 1


worksheet.write('E16', 0)
#==========================================================
#Creates a total row for each column and writes them 

worksheet.write('A17', 'Totals: ')

worksheet.write('B17', total_entering_Throughout_day)

worksheet.write('C17', total_leaving_throughout_day)

total_season_tickets = 0
for i in SeasonTicketsInPark:
    total_season_tickets += i

worksheet.write('D17', total_season_tickets)
worksheet.write('E17', total_day_pass_holders)
#===================================================
#Creates averages for season ticket holders and day pass holders and writes them

average_season_ticket_holders = total_season_tickets / 14
worksheet.write('D18', average_season_ticket_holders)

average_day_pass_holders = total_day_pass_holders / 14
worksheet.write('E18', average_day_pass_holders)




#worksheet.autofit()
workbook.close()