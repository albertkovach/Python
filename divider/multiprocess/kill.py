#kill a Process: Chapter 3: Process Based Parallelism
import multiprocessing
import time

def foo():
    print ('Starting function')
    time.sleep(0.1)
    print ('Finished function')

if __name__ == '__main__':
    p = multiprocessing.Process(target=foo)
    p.daemon = False
    print ('Process before execution:', p, p.is_alive())
    p.start()
    print("")
    print ('Process running:', p, p.is_alive())
    p.terminate()
    print("")
    print ('Process terminated:', p, p.is_alive())
    p.join()
    print("")
    print ('Process joined:', p, p.is_alive())
    print ('Process exit code:', p.exitcode)