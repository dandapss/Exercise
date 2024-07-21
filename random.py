import random

states = ["Delaware", "Los Angelas", "New York"]

numbers = len(states)
print(numbers)
random_number = random.randint(0, numbers - 1)
print(random_number)
bla = states[random_number]

print(f"Move to {bla}, there is where you need to live")