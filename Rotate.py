# Rotate
the_input = "□□■□□\n□■■■□\n■□■□■\n□□■□□\n□□■□□"

list = []
allow = the_input.split()
for each in allow:
    list2 = []
    for square in each:
        list2.append(square)
    list.append(list2)

lengg = len(list[0])
for first in list:
    place = list.index(first)
    first = list[lengg-place]
print(list)


# not done.. still in progress
