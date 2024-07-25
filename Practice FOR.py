# find height average (must use "for")

# student_height = input("Tell me the people's height: ").split()
# number = len(student_height)
# sum_of_height = 0

# for tt in range(0, number):
#     sum_of_height+=int(student_height[tt])

# average = sum_of_height/number

# print(average)

# find the highest score without using "max()" but "for"

student_score = input().split()
highest = 0
for ttt in student_score:
    ttt = int(ttt)
    if ttt > highest:
        highest = ttt
print(ttt)