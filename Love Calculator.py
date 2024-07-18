# Love Calculator

print("Welcome to the Love Calculator\n Type your name and partner's to check the score")
name1 = input("Your name: ")
name2 = input("Your partner's name: ")
fullname = name1 + " " + name2
print(fullname)
Upper_case = fullname.upper()
name1_T = Upper_case.count("T")
print(f"T occurs {name1_T} time(s)")
name1_R = Upper_case.count("R")
print(f"R occurs {name1_R} time(s)")
name1_U = Upper_case.count("U")
print(f"U occurs {name1_U} time(s)")
name1_E = Upper_case.count("E")
print(f"E occurs {name1_E} time(s)")

name1_total = int(name1_T) + int(name1_R) + int(name1_U) + int(name1_E)
print(f"   The TRUE score is {name1_total}")

name2_L = Upper_case.count("L")
print(f"L occurs {name2_L} time(s)")
name2_O = Upper_case.count("O")
print(f"O occurs {name2_O} time(s)")
name2_V = Upper_case.count("V")
print(f"V occurs {name2_V} time(s)")
name2_E = Upper_case.count("E")
print(f"E occurs {name2_E} time(s)")

name2_total = int(name2_L) + int(name2_O) + int(name2_V) + int(name2_E)
print(f"   The LOVE score is {name2_total}")

TT = name1_total*10
Score = TT + name2_total

if Score < 10 or Score > 90:
    print(f"\nYour score is {Score}, you go together like coke and mentos")
elif Score >=40 and Score <= 50:
    print(f"\nYour score is {Score}, you are alright together")
else:
    print("\n#######################")
    print(f"Your total score is {Score}")
    print("#######################")

# random Love Score less than certain number
import random

random_number = random.random() * 5
print(f"Your love score is {random_number}")


# Treasure Island game

print("Welcome to Treasure Island.\nYour mission is to find the treasure.\n")
leftorright = input("Choose 'left' or 'right': ")
if leftorright == "right":
    print("Game Over")
elif leftorright == "left":
    print("You survived!!")
    swimorwait = input("Choose 'swim' or 'wait': ")
    if swimorwait == "swim":
        False
    elif swimorwait == "wait":
        print("You survived!!")
        door = input("Choose 'red' or 'blue' or 'yellow': ")
        if door == "red":
            False
        elif door == "blue":
            False
        elif door == "yellow":
            print("You Win!!!!!")
