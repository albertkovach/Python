def ProcKiller2(path)
    
    for p in psutil.process_iter()
        print(p.name())
        try
            if path in str(p.open_files())
                print(p.name())
                print(^^^^^^^^^^^^^^^^^)
                p.kill()
        except
            continue


def ProcKiller(path)
    print(^^^^^^ ProcKiller 1 ^^^^^^^^^)
    pathToClear = path
    print('pathToClear {0}'.format(pathToClear))

    print(f...Searching for active processes matching path '{pathToClear}')
    for proc in psutil.process_iter(['name', 'open_files'])
        for file in proc.info['open_files'] or []
            print(file.path)
            #if pathToClear in file.path
            #    print(fFound process potentiall process potentially locking path. Killing '{proc.info['name']}')
            #    print(%-5s %-10s %s % (proc.pid, proc.info['name'][10], file.path))
            #    proc.kill()

    print(Completed!)