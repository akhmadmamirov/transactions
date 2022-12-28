#pip install openpyxl
import openpyxl

# Open the Excel file
wb = openpyxl.load_workbook('Final_Intern-Account_Transactions.xlsx')

# Get a sheet by name Data from the Excel File
sheet = wb['Data']


#Initiating User Database
users = {}

#Iterating through each user
for i in range(2, 1000):
    row = sheet[i]
    #Getting Username
    user = row[0].value
    #Getting Amount
    amount = row[2].value
    if user not in users and user is not None:
        #Creating a user with [Ending Balance, Maximum Balance, Minimum Balance]
        users[user] = [0,float("-inf"),float("inf")]
    if user is not None:
        #ending balance
        users[user][0] +=int(amount)
        #Maximum balance
        users[user][1] = max(users[user][0],users[user][1])
        #Minimum balance
        users[user][2] = min(users[user][0],users[user][2])

# Print user values [Ending Balance, Maximum Balance, Minimum Balance]
print(users)



# Create a new Excel file
wb2 = openpyxl.Workbook()

# Get the active sheet
sheet2 = wb2.active
#Adding Column Header
data = [
    ['Users', 'Ending Balance', 'Maximum Balance', 'Minimum Balance']
    ]

for row in data:
    sheet2.append(row)

#Adding Data from ou Users Database
for cell, values in users.items():
    sheet2.append([cell] + values)

wb2.save('results.xlsx')