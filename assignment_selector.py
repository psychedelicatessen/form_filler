prompt = ">"

# a function for user assignment choice
def assign_select(assignments):
    
    count = 1

    print('Choose an assignment.')
    for num in range(0, len(assignments)):
        print(str(count) + ". " + assignments[num]['Name'])
        count += 1
    # a loop which gets the user's target assignment for input
    while True:
        try:
            target = int(input(prompt))
        except ValueError:
            print('Enter a number between 1 and ' + str(count))
            continue
        if target > count:
            print('Your choice can\'t be greater than ' + str(count) +'!')
        else:
            break

    # set the target as the chosen assignment dictionary
    target = assignments[target-1]

    return target

# a function for user transport choice
def trans_select(transports):
    
    count = 1

    print('Choose transport.')
    for num in range(0, len(transports)):
        print(str(count) + ". " + transports[num]['Name'])
        count += 1
    # a loop which gets the user's target assignment for input
    while True:
        try:
            target = int(input(prompt))
        except ValueError:
            print('Enter a number between 1 and ' + str(count))
            continue
        if target > count:
            print('Your choice can\'t be greater than ' + str(count) +'!')
        else:
            break

    # set the target as the chosen transport dictionary
    target = transports[target-1]

    return target
