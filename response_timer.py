import time

#start timer 
def start_timer():
    return time.perf_counter()

#return elapsed time 
def stop_timer(start_time):

    elapsed = time.perf_counter() - start_time
    return round(elapsed, 2)