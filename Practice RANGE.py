# sum of even numbers between 0 and certain number
number = int(input("What is your number?: "))
t1=0
for t in range(2, number+1):
    if t%2 >= 1:
        print("odd", t)
    else:
        t1+=t
        print("even", t)
print(t1)
######## or
# number = int(input())
# t1=0
# for t in range(2, number+1, 2):
#     t1+=t

# create fizzbuzz (Korean 369 game including number 5)
print("Rule No.1 : If the number is divisible by 3, call 'FIZZ' instead of the number.")
print("Rule No.2 : If the number is divisible by 5, call 'BUZZ' instead of the number.")
print("Rule No.3 : If the nubmer is divisible by both 3 and 5 call 'FIZZBUZZ'.")

numbers = int(input("What are your numbers?: "))
fizz = 0
for number in range(1, numbers+1):
    if number%3 == 0 and number%5 == 0:
        print("FIZZBUZZ")
    elif number%3 == 0:
        print("FIZZ")
    elif number%5 == 0:
        print("BUZZ")
    else:
        print(number)
