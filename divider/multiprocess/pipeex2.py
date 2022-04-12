from multiprocessing import Pipe, Process

def proc_a(pipe):
    data = 25
    pipe.send(data)
    result = pipe.recv()
    print(result)

def proc_b(pipe):
    data = pipe.recv()
    data_new = data + 1 # Sample change 
    pipe.send(data_new)

if __name__ == '__main__':
    parent, child = Pipe()
    p1 = Process(target=proc_a, args=(parent,))
    p2 = Process(target=proc_b, args=(child,))

    p1.start()
    p2.start()
    p1.join()
    p2.join()
    
    # parent <======> child
    # proc_a <======> proc_b