import random

###############################################
# List
word_list = ['test', 'hello', 'ardvark']
# random word from the list
word = random.choice(word_list)
print(f"Your word is {word}")
# find length of the random word to find right place of the letter
length = len(word)
# make list of "_" as length of the random word
display = []
for letter in range(length):
    display+="_"
# death count of the game
Mode = input("Choose your mode. Type 'EASY' or 'HARD' \n")
mode = Mode.lower()
if mode == "easy":
    wrong = length
elif mode == "hard":
    wrong_list = []
    for letter in word:
        if letter not in wrong_list:
            wrong_list.append(letter)
    wrong = len(wrong_list)
else:
    print("I gave only two options, but you choose neither. Game lose by not following the rule")
    exit()
    
print(f"This is your death count: {wrong}")
print(display)
###############################################

while wrong > 0 and "_" in display:
    choice = input("What is your letter?: ")
    if choice in word and choice not in display:
        for position in range(length):
            each_letter = word[position]
            if each_letter == choice:
                display[position] = choice
        print(display)
    elif choice in display:
        print("Do not repeat the same letter, negative 1 point for dumbness")
        wrong-=1
        print(f"Your remaining life is {wrong}")
    else:
        wrong-=1
        print(f"Your remaining life is {wrong}")

if "_" not in display:
    print("You saved your man")
elif wrong == 0:
    print("Your man died, game lose")

# 24.07.29
# Things to add later on
# 1. If the user repeat the same letter, give negative point -- done
# 2. Count only once of the repeated letter in the death count. -- done
