import time
import sys
import colorama
colorama.init()

print('line 1')
print('line 2')
print('line 3')
print('line 4')
print('line 5')
print('line =')
print('line =')

for i in range(500):
    #print('{0} из 200000'.format(i), end='\r')
    
    sys.stdout.write(' \r\033[K')
    sys.stdout.write('\033[1A') # Up a line
    sys.stdout.write(' \r\033[K')
    sys.stdout.write('\033[1A') # Up a line
    sys.stdout.write(' \r\033[K') # Delete current line
    print('first {0} из 200000'.format(i))
    print('second {0} из 200000'.format(i))
    


    #sys.stdout.write('\033[1A') # Up a line
    #sys.stdout.write(' \r\033[K') # Delete current line

    
    
    sys.stdout.flush()
    time.sleep(0.5)