from datetime import datetime
import os

date= os.path.getmtime('test.xlsm')
real_time= datetime.fromtimestamp(date)
print(date)
print(datetime.now())
print(real_time)
if real_time > datetime.now():
    print('real\'s higher')
else:
    print('now is higher')