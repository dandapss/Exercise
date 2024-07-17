print("Welcome to the tip calculator.")
bill = float(input("What was the total bill?: $"))
tip = float(input("What percentage tip would you like to give? 10,12, or 15?: ")) / 100
percentage = float(1 + tip)
whole_price = bill * percentage
people = float(input("How many people to split the bill?: "))
each = whole_price / people
final_amount = "{:.2f}".format(each)
print(f"Each person should pay: ${final_amount}")






# tip = float(input("What percentage tip would you like to give? 10,12, or 15?: ")) / 100
# whole_price = tip + bill
# people = float(input("How many people to split the bill?: "))
# each = round((whole_price / people),2)
# print(f"Each person should pay: ${each}")