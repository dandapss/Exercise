# Rotate
the_input = "□□■□□\n□■■■□\n■□■□■\n□□■□□\n□□■□□"

list = []
allow = the_input.split()
for each in allow:
    list2 = []
    for square in each:
        list2.append(square)
    list.append(list2)

length = len(list[0])
for first_list in list:
    place = list.index(first_list)
    first_list = list[length - place]
    
print(list)




# not done.. still in progress
