import pandas as pd
import numpy as np
import os

excel_filer = [file for file in os.listdir() if file.endswith('.xlsx')]
if len(excel_filer) == 0:
    print("No Excel files found in the directory.")
    exit(0)

first = excel_filer[0]
data = pd.read_excel(first)

pre = data[
        (data['Pre-Installed'] == True) & (data['Requires Assistance'] == False)
        ]
ava = data[
        (data['Pre-Installed'] == False) & (data['Distribution Type'] == 'Software Center (Windows)')
        ]
prenavn = np.array(pre['Name'])
avnavn = np.array(ava['Name'])

print("Hei CUSTOMER\n")
print("Vi starter nå med oppsettet av maskinen til XXX\n")
print("Følgende applikasjoner vil bli installert på PC-en av Intility:")

for i, app in enumerate(prenavn, 1):
    print(f"{i}. {app}")

print("\nFølgende applikasjoner vil kun ligge tilgjengelig for installasjon i Software Center")

for i, app in enumerate(avnavn, 1):
    print(f"{i}. {app}")

print("\nGi beskjed dersom noe ikke stemmer.")
print("\nVi gir beskjed når maskinen er ferdigstilt, og er klar til å sendes ut. PC-en og eventuelt annet ekstrautstyr vil bli sendt til følgende adresse:XXXXXXXXXXXXXXXXXXXXXXXXX")
