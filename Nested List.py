# Nested List
dirty_dozen = ["Strawberries", "Spinach", "Kale", "Nectarines", "Apples", "Grapes", "Peaches", "Cherries", "Pears", "Tomatoes", "Celery", "Potatoes"]

fruits = ["Strawberries", "Nectarines", "Apples", "Grapes", "Peaches", "Cherries", "Pears"]
vegetables = ["Spinach", "Kale", "Tomatoes", "Celery", "Potatoes"]

dozen = [fruits, vegetables]
print(dozen)
print(dozen[0][1])


# Hiding treasure map game
line1 = ["ㅁ", "ㅁ", "ㅁ"]
line2 = ["ㅁ", "ㅁ", "ㅁ"]
line3 = ["ㅁ", "ㅁ", "ㅁ"]
map = [line1, line2, line3]
print("Hiding your treasure! X marks the spot.")
position = input("Where do you want to hide the treasure?!(ex. a1): ")
letter = position[0].upper()
number = int(position[1])

if number == 1:
    if letter == "A":
        line1[0] = "X"
    elif letter == "B":
        line1[1] = "X"
    elif letter == "C":
        line1[2] = "X"
    else:
        print(f"Wrong position {position}")
if number == 2:
    if letter == "A":
        line2[0] = "X"
    elif letter == "B":
        line2[1] = "X"
    elif letter == "C":
        line2[2] = "X"
    else:
        print(f"Wrong position {position}")
if number == 3:
    if letter == "A":
        line3[0] = "X"
    elif letter == "B":
        line3[1] = "X"
    elif letter == "C":
        line3[2] = "X"
    else:
        print(f"Wrong position {position}")

####### easier way
# abc = ["A", "B", "C"]
# letter_index = abc.index(letter)
# number_index = int(number - 1)
# map[number_index][letter_index] = "X"

print(f"{line1}\n{line2}\n{line3}")
