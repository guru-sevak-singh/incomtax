from datetime import datetime
import time

starting_time = datetime.now()

print(starting_time)

input('Enter any Thing...')

first_lap = datetime.now()

first_gap = first_lap - starting_time
print(first_gap.seconds)


if first_gap.seconds < 60:
    print('chota h')

else:
    print('bada h....')