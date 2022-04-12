import multiprocessing
import time

def test():
    print("start test process")
    time.sleep(1)
    print("end test process")


if __name__ == "__main__":
    process = multiprocessing.Process(target=test)
    print(f"The process is alive:{process.is_alive()}")
    process.start()
    print(f"The process is alive:{process.is_alive()}")
    process.join()
    print(f"The process is alive:{process.is_alive()}")