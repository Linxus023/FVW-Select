### Importeren van packages voor Python Script

import os
import regex as re
import pandas as pd

### Locatie aanwijzen voor huidige working directory
cwd = os.getcwd()

### Load data voor gebruikers & acties
users = pd.read_excel(cwd + "/users.xlsx")
data = pd.read_excel(cwd + "/data.xlsx")

### Hier geef ik aan wat ik zoek in de Excel lijst (FVW is een urenregistratie programma)
regex = "FVW"

### Aanmaken Boolean voor input
booleanYesNo = input("Wil jij je actie lijst zien (Y/N)?\n")

### While functie zorgt er voor dat gebruiker in het script blijft of accepteert Y
while booleanYesNo != "Y":
    if booleanYesNo == "N":
        print("Antwoord Y graag....\n")
    else:
        print("Helaas, doe Y of N\n") 
    booleanYesNo = input("Wil jij je actie lijst zien?\n")

### Print een overview van de huidige meldingen in de database
print(data['Acties'])
print("-"*40)

### Vraagt naar de UserID om te identificeren wie er inlogt
isUserID = input("Wat is je userID?\n")
isUserID = int(isUserID)
userList = users['userID'].tolist()

### While functie om te checken of UserID wel bestaat in de database
while isUserID not in userList:
    try:
        print("Deze userID bestaat niet volgens onze Database/n")
        isUserID = input("Wat is je userID?\n")
        isUserID = int(isUserID)
    except:
        print("Deze userID bestaat niet volgens onze Data Base/n")
        isUserID = input("Wat is je userID?\n")
        isUserID = int(isUserID)

### Print de gekozen User
print("-"*40)
logUserName = users.loc[users['userID'] == isUserID]
logUserName = logUserName['userName'].item()

print("Welkom " + str(logUserName))
print("-"*40)
print("De FWV acties worden op jouw naam gezet!")

### Set de acties met keyword "FVW" op naam van gebruiker
def registratieOpNaam(row):
    resultaat = re.match(regex, row['Acties'])
    if resultaat is not None:
        return row['loggedUser']
    else:
        return row['OplosgroepSorteren']

### Print de nieuwe dataframe en exporteer deze naar excel
data['loggedUser'] = logUserName
data['OplosgroepSorterenNieuw'] = data.apply(registratieOpNaam, axis=1)
print(data[['Acties', 'OplosgroepSorteren', 'OplosgroepSorterenNieuw']].to_string())
print("-"*40)
print("De nieuwe excel met acties op jouw naam staat klaar in de folder!")

data.to_excel(cwd + "/data_new.xlsx", index=False)

exit()



        

