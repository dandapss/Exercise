# Def, For, While practice
# used <https://reeborg.ca/reeborg.html?lang=en&mode=python&menu=worlds%2Fmenus%2Freeborg_intro_en.json&name=Hurdle%204&url=worlds%2Ftutorial_en%2Fhurdle4.json> 

# Hurdle No.3
def jump():
    turn_left()
    move()
    turn_right()
    move()
    turn_right()
    move()
    turn_left()
    
def turn_right():
    turn_left()
    turn_left()
    turn_left()
def turn_left_move():
    turn_left()
    move()

while not at_goal():
    if front_is_clear():
        move()
    elif wall_in_front():
        jump()
            
            
# Hurdle No.4
def turn_right():
    turn_left()
    turn_left()
    turn_left()
def turn_left_move():
    turn_left()
    move()
    
count = 0
while not at_goal():
    if wall_in_front() and not is_facing_north():
        turn_left()
        while wall_on_right() and not wall_in_front() and not at_goal():
            move()
    elif is_facing_north():
        turn_right()
        move()
        turn_right()
    else:
        move()

# Maze
def turn_right():
    turn_left()
    turn_left()
    turn_left()
def find_way():
    if right_is_clear():
        turn_right()
        move()
    elif front_is_clear():
        move()
    else:
        turn_left()
            
        
while not at_goal():
    find_way()
    
