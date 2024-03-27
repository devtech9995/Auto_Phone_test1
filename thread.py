import time
import threading

def calc_square(numbers):
    for n in numbers:
        print(f'\n{n} ^ 2 = {n*n}')
        time.sleep(0.1)

def calc_cube(numbers):
    for n in numbers:
        print(f'\n{n} ^ 3 = {n*n*n}')
        time.sleep(0.1)

numbers = [2, 3, 5, 8]
# start = time.time()

##### Implementation of two functions #####
# calc_square(numbers)
# calc_cube(numbers)
###########################################

############# Multithreading ##############
square_thread = threading.Thread(target=calc_square, args=(numbers,))
cube_thread = threading.Thread(target=calc_cube, args=(numbers,))

square_thread.start()
cube_thread.start()

square_thread.join()
cube_thread.join()
###########################################

# end = time.time()

print('Execution Time: {}'.format(end-start))