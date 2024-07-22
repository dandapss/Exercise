import random

print("Welcome to Rock Paper and Scissors game!\n")
my_choice = int(input("What do you choose? Type 0 for Rock, 1 for Paper or 2 for Scissors: "))
Computer = random.randint(0, 2)
result = my_choice - Computer

if my_choice >= 0 and my_choice <= 3:
    # Tried to use "def"
    def mchoice():
        if my_choice == 0:
            mchoice = "Rock"
        elif my_choice == 1:
            mchoice = "Paper"
        elif my_choice == 2:
            mchoice = "Scissors"   
        return mchoice

    def cchoice():
        if Computer == 0:
            cchoice = "Rock"
        elif Computer == 1:
            cchoice = "Paper"
        elif Computer == 2:
            cchoice = "Scissors"
        return cchoice
        
    choice1 = mchoice()
    choice2 = cchoice()
    print(f"you got {choice1} and Computer got {choice2}")

# other way
    # index = ["Rock", "Paper", "Scissors"]
    # M_index = index[my_choice]
    # C_index = index[Computer]
    # print(f"You got {M_index} and computer got {C_index}")


    if my_choice == Computer:
        print("It's a draw!")
    elif result == -1 or result == 2:
        print("You Lose!")
    else:
        print("You Won!")
else:
    print("That is not valid number!!!")
