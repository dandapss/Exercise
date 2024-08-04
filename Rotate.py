# Rotate
the_input = "□□■□□\n□■■■□\n■□■□■\n□□■□□\n□□■□□"

input_list = []
allow = the_input.split()
for each in allow:
    input_list.append(list(each))

length = len(list[0])
def rotate(tt):
    new_list = []
    for x in range(length):
        temp_list = []
        for y in range(length):
            temp_list.append(tt[length - 1 - y][x])
        new_list.append(temp_list)
    return new_list

def degree(tt):
    for t in tt:
        print(''.join(t))

answer = int(input("How many rotations do you want?: "))
one = rotate(input_list)
two = rotate(one)
three = rotate(two)
if answer == 1:
    degree(one)
elif answer == 2:
    degree(two)
elif answer == 3:
    degree(three)
elif answer == 0:
    degree(input_list)
else:
    print("Only 1, 2, and 3 are developed")

# not done.. still in progress -- done(24.08.04)

