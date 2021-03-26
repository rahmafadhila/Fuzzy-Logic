import pandas as pd
import xlsxwriter

#FUZZIFIKASI PENGHASILAN
def Penghasilan(x):
	def PenghasilanUpper(x):
		if (x > 16):
			return 1
		elif ((x > 13) and (x <= 16)):
			return ((x - 13) / (16 - 13))
		elif (x <= 13):
			return 0

	def PenghasilanMiddle(x):
		if ((x > 12) and (x <= 16)):
			return ((16 - x) / (16 - 12))
		elif ((x > 9) and (x <= 12)):
			return 1
		elif ((x > 5) and (x <= 9)):
			return ((x - 5) / (9 - 5))
		elif ((x <= 5) or (x > 16)):
			return 0

	def PenghasilanBottom(x):
		if (x <= 5):
			return 1
		elif ((x > 5) and (x <= 8)):
			return ((8 - x) / (8 - 5)) 
		elif (x > 8):
			return 0

	IncomeUpper = PenghasilanUpper(x)
	IncomeMiddle = PenghasilanMiddle(x)
	IncomeBottom = PenghasilanBottom(x)
	return IncomeUpper, IncomeMiddle, IncomeBottom

#FUZZIFIKASI PENGELUARAN
def Pengeluaran(x):
	def PengeluaranUpper(x):
		if (x > 9):
			return 1
		elif ((x > 7) and (x <= 9)):
			return ((x - 7) / (9 - 7))
		elif (x <= 7):
			return 0

	def PengeluaranMiddle(x):
		if ((x > 6) and (x <= 9)):
			return ((9 - x) / (9 - 6))
		elif ((x > 3) and (x <= 6)):
			return 1
		elif ((x > 0) and (x <= 3)):
			return ((x - 0) / (3 - 0))
		elif ((x <= 0) or (x > 9)):
			return 0

	def PengeluaranBottom(x):
		if (x <= 0):
			return 1
		elif ((x > 0) and (x <= 2)):
			return ((2 - x) / (2 - 0)) 
		elif (x > 2):
			return 0

	SpendingUpper = PengeluaranUpper(x)
	SpendingMiddle = PengeluaranMiddle(x)
	SpendingBottom = PengeluaranBottom(x)
	return SpendingUpper, SpendingMiddle, SpendingBottom

#------------FUZZY RULES INFERENSI-----------#
#| PENGHASILAN |  PENGELUARAN  |    SCORE   |
#| ----------- |  -----------  |   -------  |
#|    UPPER    |     UPPER     | REJECTED   |
#|    UPPER    |     MIDDLE    | REJECTED   |
#|    UPPER    |     BOTTOM    | REJECTED   |
#|    MIDDLE   |     UPPER     | ACCEPTED   |
#|    MIDDLE   |     MIDDLE    | CONSIDERED |
#|    MIDDLE   |     BOTTOM    | CONSIDERED |
#|    BOTTOM   |     UPPER     | ACCEPTED   |
#|    BOTTOM   |     MIDDLE    | ACCEPTED   |
#|    BOTTOM   |     BOTTOM    | ACCEPTED   |

#INFERENSI
def Inferensi(IncomeUpper, IncomeMiddle, IncomeBottom, SpendingUpper, SpendingMiddle, SpendingBottom): 
	#MENCARI NILAI SCORE
	FuzzyRules =	[[min(IncomeUpper, SpendingUpper), 'REJECTED'],
	  				[min(IncomeUpper, SpendingMiddle), 'REJECTED'],
	  				[min(IncomeUpper, SpendingBottom), 'REJECTED'],
	  				[min(IncomeMiddle, SpendingUpper), 'ACCEPTED'],
	  				[min(IncomeMiddle, SpendingMiddle), 'CONSIDERED'],
	  				[min(IncomeMiddle, SpendingBottom), 'CONSIDERED'],
	  				[min(IncomeBottom, SpendingUpper), 'ACCEPTED'],
	  				[min(IncomeBottom, SpendingMiddle), 'ACCEPTED'],
	  				[min(IncomeBottom, SpendingBottom), 'ACCEPTED']]

	allRejected = []
	allConsidered = []
	allAccepted = []

	for i in range(len(FuzzyRules)):
	  	if FuzzyRules[i][1] == 'ACCEPTED':
	  		allAccepted.append(FuzzyRules[i][0])
	  	elif FuzzyRules[i][1] == 'CONSIDERED':
	  		allConsidered.append(FuzzyRules[i][0])
	  	elif FuzzyRules[i][1] == 'REJECTED':
	  		allRejected.append(FuzzyRules[i][0])

	#NILAI MAX MASING-MASING SCORE
	return max(allRejected), max(allConsidered), max(allAccepted) 

#DEFUZZIFIKASI SUGENO
def Defuzzifikasi(allRejected, allConsidered, allAccepted):
  a = ((allRejected * 50) + (allConsidered * 70) + (allAccepted * 100))
  b = (allRejected + allConsidered + allAccepted)
  c = a / b
  return round(c, 2)


#MAIN PROGRAM#
#READFILE
data = pd.read_excel('Mahasiswa.xls', 'Mahasiswa')
index = (data['Id']).values.tolist()
penghasilan = (data['Penghasilan']).values.tolist()
pengeluaran = (data['Pengeluaran']).values.tolist()

bantuan = []
for i in range(len(index)):
  IncomeUpper, IncomeMiddle, IncomeBottom = Penghasilan(penghasilan[i])
  SpendingUpper, SpendingMiddle, SpendingBottom = Pengeluaran(pengeluaran[i])
  allRejected, allConsidered, allAccepted = Inferensi(IncomeUpper, IncomeMiddle, IncomeBottom, SpendingUpper, SpendingMiddle, SpendingBottom)
  Score = Defuzzifikasi(allRejected, allConsidered, allAccepted)
  bantuan.append(Score)

#SORTING SCORE DEFUZZIFIKASI
for i in range(len(index)-1, 0, -1):
	max = 0
	for j in range(1, i+1):
		premax = index[max]
		if bantuan[j] < bantuan[max]:
			max = j         
		itemp = index[i]
		index[i] = index[max]
		index[max] = itemp
			
		ptemp = penghasilan[i]
		penghasilan[i] = penghasilan[max]
		penghasilan[max] = ptemp

		qtemp = pengeluaran[i]
		pengeluaran[i] = pengeluaran[max]
		pengeluaran[max] = qtemp

		btemp = bantuan[i]
		bantuan[i] = bantuan[max]
		bantuan[max] = btemp

#WRITEFILE
dataFinal = []
for i in range(0, 20):
	final = index[i]
	dataFinal.append(final)

df = pd.DataFrame({'Id' : dataFinal})
write = pd.ExcelWriter('Bantuan.xlsx', engine = 'xlsxwriter')
df.to_excel(write, sheet_name = 'Bantuan', index = False)
write.save()