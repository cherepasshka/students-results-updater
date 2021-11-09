import os

res_made = os.system('python3 maketable.py > obj.txt')
if res_made != 0:
	print('Something went wrong while generating data')
	exit(1)
sheet_built = os.system('python3 main.py')

