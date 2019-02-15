import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore
import xlsxwriter

# Use a service account
cred = credentials.Certificate('C:/Users/HARI VARUN/Documents/comsurvey-95954-8691c53b39c3.json')
firebase_admin.initialize_app(cred)

db = firestore.client()

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('AnantapurUrban.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Write some simple text.
worksheet.write('A1', 'ID',bold)
worksheet.write('B1', 'LOCATION',bold)
worksheet.write('C1', 'DATE',bold)
worksheet.write('D1', 'MANDAL',bold)
worksheet.write('E1', 'PANCHAYAT',bold)
worksheet.write('F1', 'VILLAGE',bold)
worksheet.write('G1', 'NAME',bold)
worksheet.write('H1', 'GENDER',bold)
worksheet.write('I1', 'AGE',bold)
worksheet.write('J1', 'CASTE',bold)
worksheet.write('K1', 'RELIGION',bold)
worksheet.write('L1', 'SUBCASTE',bold)
worksheet.write('M1', 'EDUCATION',bold)
worksheet.write('N1', 'OCCUPATION',bold)
worksheet.write('O1', 'FINANCIAL STATUS',bold)
worksheet.write('P1', 'PHONE NUMBER',bold)
worksheet.write('Q1', 'Q1',bold)
worksheet.write('R1', 'Q2',bold)
worksheet.write('S1', 'Q3',bold)
worksheet.write('T1', 'Q4',bold)
worksheet.write('U1', 'Q5',bold)
worksheet.write('V1', 'Q6',bold)
worksheet.write('W1', 'Q7',bold)
worksheet.write('X1', 'Q8',bold)
worksheet.write('Y1', 'Q9',bold)
worksheet.write('Z1', 'Q10',bold)
worksheet.write('AA1', 'Q11',bold)
worksheet.write('AB1', 'Q12',bold)
worksheet.write('AC1', 'Q13',bold)
worksheet.write('AD1', 'Q14',bold)
worksheet.write('AE1', 'Q15',bold)
worksheet.write('AF1', 'Q16',bold)
worksheet.write('AG1','SURVEY PERSON ID',bold)



# Write some numbers, with row/column notation.
worksheet.write(2, 0, 123)
worksheet.write(3, 0, 123.456)

docs = db.collection(u'AnanthapurUrban').get()
count = 1
for doc in docs:
    p = doc.to_dict()
    worksheet.write(count,0,doc.id)
    worksheet.write(count,1,p['Location'])
    worksheet.write(count,2,p['Date'])
    worksheet.write(count,3,p['Mandal'])
    worksheet.write(count,4,p['Panchayat'])
    worksheet.write(count,5,p['Village'])
    worksheet.write(count,6,p['Name'])
    worksheet.write(count,7,p['Gender'])
    worksheet.write(count,8,p['Age'])
    worksheet.write(count,9,p['Caste'])
    worksheet.write(count,10,p['Religion'])
    worksheet.write(count,11,p['Subcaste'])
    worksheet.write(count,12,p['Education'])
    worksheet.write(count,13,p['Occupation'])
    worksheet.write(count,14,p['FinancialStatus'])
    worksheet.write(count,15,p['Phone'])
    worksheet.write(count,16,p['Q1'])
    worksheet.write(count,17,p['Q2'])
    worksheet.write(count,18,p['Q3'])
    worksheet.write(count,19,p['Q4'])
    worksheet.write(count,20,p['Q5'])
    worksheet.write(count,21,p['Q6'])
    worksheet.write(count,22,p['Q7'])
    worksheet.write(count,23,p['Q8'])
    worksheet.write(count,24,p['Q9'])
    worksheet.write(count,25,p['Q10'])
    worksheet.write(count,26,p['Q11'])
    worksheet.write(count,27,p['Q12'])
    worksheet.write(count,28,p['Q13'])
    worksheet.write(count,29,p['Q14'])
    worksheet.write(count,30,p['Q15'])
    worksheet.write(count,31,p['Q16'])
    worksheet.write(count,32,p['SurveyPersonName'])
    count = count + 1 
workbook.close()