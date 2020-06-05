import datetime
import time

d = datetime.datetime(1991, 9, 3)#+datetime.timedelta(seconds=25)
print(d)

e = datetime.datetime(1991, 9, 2)
print(e)


print(e.timestamp()-d.timestamp())