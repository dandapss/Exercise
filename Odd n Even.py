# Pratice No.1 : Use multiple "if"
height = int(input())
prise = 0

if height > 120:
    print("Your can ride it!!!")
    age = int(input("What is your age?: "))
    if age >= 18:
        ticket = prise+18
    if age < 12:
        ticket = prise+5
    if age >= 12 and age < 18:
        ticket = prise+7
    Photo = input("Do you want photo?: ")
    if Photo == "Yes":
        total = ticket + 3
        print(f"The total bill is {total}")
    else:
        total = ticket
        print(f"The total bill is {total}")
else:
    print("Grow more short man")


# Pratice No.2 : Use multiple "if"
print("Thank you for choosing Pizza Deliveries!!")
size = input("Choose the size of the pizza: (L, M, or S) ")
prize = 0
if size == "L":
    prize = int(15)
    print(f"The prize is: ${prize}")
if size == "M":
    prize = int(12)
    print(f"The prize is: ${prize}")
if size == "S":
    prize = int(10)
    print(f"The prize is: ${prize}")

add_pepperoni = input("Would you like to add some pepperoni?: (Y or N) ")
if add_pepperoni == "Y":
    prize+=5
print(f"The prize is: ${prize}")

add_cheese = input("Would you like some cheese?: (Y or N)")
if add_cheese == "Y":
    prize+=1

print(f"The total prize is: ${prize}")

