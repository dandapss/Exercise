# Nested List
dirty_dozen = ["Strawberries", "Spinach", "Kale", "Nectarines", "Apples", "Grapes", "Peaches", "Cherries", "Pears", "Tomatoes", "Celery", "Potatoes"]

fruits = ["Strawberries", "Nectarines", "Apples", "Grapes", "Peaches", "Cherries", "Pears"]
vegetables = ["Spinach", "Kale", "Tomatoes", "Celery", "Potatoes"]

dozen = [fruits, vegetables]
print(dozen)

# If there are two or more lists in a list, the first [] means which list to select and the second one means which one of the list. (Of course, both of them start from 0)
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
# I have to type one of A, B, or C so make list of them
# abc = ["A", "B", "C"]
# Check which of the letter I typed and where of the list is in
# letter_index = abc.index(letter)
# The index starts from 1 so subtract 1 from the index
# number_index = int(number - 1)
# Use nested list to change info in list.
# map[number_index][letter_index] = "X"

print(f"{line1}\n{line2}\n{line3}")
