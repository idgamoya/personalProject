from datetime import datetime

now = datetime.now()
wfname = 'result' + now.strftime('%Y%m%d%H%M%S') + '.xlsx'
print(wfname)
