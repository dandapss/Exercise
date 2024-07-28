# Password Generator Project

import random

password = []
letters = ["a", "b", "c", "d"]
numbers = ['0', '1', '2', '3']
symbols = ['!', '@', '#', '$']

print("Welcome to the Python Password Generator!!")
my_letters = int(input("How many letters would you like in your password?: "))
my_numbers = int(input("How many numbers would you like in your password?: "))
my_symbols = int(input("How many symbols would you like in your password?: "))

for tt in range(0,my_letters):
    ran = random.randint(0,3)
    print(ran)
    let = letters[ran]
    password.append(let)
    print(password)

for tt in range(0,my_numbers):
    ran = random.randint(0,3)
    num = numbers[ran]
    password.append(num)
    print(password)

for tt in range(0,my_symbols):
    ran = random.randint(0,3)
    sym = symbols[ran]
    password.append(sym)
    print(password)

e = ''.join(password)
print(f"Easy level password: {e}")
random.shuffle(password)
h = ''.join(password)
print(f"Hard level password: {h}")

####### other way
# Hard Level
password_list = []
for char in range(1, my_letters + 1):
    password_list.append(random.choice(letters))

for char in range(1, my_numbers):
    password_list.append(random.choice(numbers))

for char in range(1, my_symbols):
    password_list.append(random.choice(symbols))

# Password in List
print(f"Normal password: {password_list}")
random.shuffle(password_list)
print(f"Hard password: {password_list}")

# Password in String
password2 = ""
for char in password_list:
    password += char

print(f"This is final password: {password}")