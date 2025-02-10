alphabet=["a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"]
print("Welcome to the encryptocami")
code=input("What is your secret code: ")
decide = input("What do you want?(encrypt or decrypte):  ")
shift = input("What is your shift number?: ")
code_list=list(code)

print(f"The code is {code_list}")
if decide == "encrypte" or decide == "enc" or decide == "e":
    print("You choose to encrypte the code")  
    new_list = []
    for letter in code_list:
        if letter in alphabet:
            Temp_Integer = alphabet.index(letter)
            Temp_Integer+=int(shift)
            print(f"The integer is {Temp_Integer}")
            if Temp_Integer < 26:
                NewLetter = alphabet[Temp_Integer]
                new_list.append(NewLetter)
            else:
                Temp_Integer-=26
                NewLetter = alphabet[Temp_Integer]
                new_list.append(NewLetter)
elif decide == "decrypte" or decide == "dec" or decide == "d":
    print("You choose to decrypte the code")
    new_list = []
    for letter in code_list:
        if letter in alphabet:
            Temp_Integer = alphabet.index(letter)
            # print(f"The integer is {Temp_Integer}")
            Temp_Integer-=int(shift)
            # print(f"The integer after the subtraction is {Temp_Integer}")
            if Temp_Integer < 0:
                NewLetter = alphabet[Temp_Integer]
                new_list.append(NewLetter)
            else:
                Temp_Integer = 26 - Temp_Integer
                NewLetter = alphabet[Temp_Integer]
                new_list.append(NewLetter)
else:
    print("Please choose either encrypte or decrypte")

e = ''.join(new_list)
print(e)


