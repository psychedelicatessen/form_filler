# command prompt
prompt = ">"

# these functions create the individual assignments

def create_assignment():

    assignment = dict()

    print("What name shall I give the assignment?")
    while True:
        name = input(prompt)
        if name != "":
            assignment['Name'] = name
            break
        else:
            print('Please enter a name!')
        
    print("What is the designation? (I/O/P/T)")
    while True:
        desgn = input(prompt)
        if desgn == 'I' or desgn == 'O' or desgn == 'P' or desgn == 'T':
            assignment['Activity'] = desgn
            break
        else:
            print('Designation may only be I/O/P/T!')

    print("What time did the assignment take place? (HH:MM - HH:MM)")
    assignment['Time'] = input(prompt)

    print("Total number of billable hours spent at the assignment.")
    while True:
        try:
            assignment['Hours'] = float(input(prompt))
            break
        except ValueError:
            print('Please enter a number.')

    print("What is your rate for the assignment?")
    while True:
        try:
            assignment['Rate'] = int(input(prompt))
            break
        except ValueError:
            print('Please enter a number.')
    
    return assignment

def create_transport():


    transport = dict()
   
    print('Input the name of the assignment.')
    while True:
        name = input(prompt)
        
        if name != "":
            
            transport['Name'] = name
            break
        
        else:
            print('Please enter a name!')
  
    print("Please input the assignment's location.")
    while True:
        loc = input(prompt)

        if loc != "":
            
            transport['Location'] = loc
            break

        else:

            print('Please enter a location!')

    print("Please input the transportation type. (TF/TX)")
    while True:
        typ = input(prompt)
        if typ == 'TF' or typ == 'TX':
                
            transport['Type'] = typ
            break
            
        else:

            print('Enter TF or TX!')

    print("Please add the number of visits.")
    while True:
        
        try:
            
            transport['Time'] = int(input(prompt))
            break
        
        except ValueError:
            
            print('Please enter a number!')

    print("Please add the cost of travel.")
    while True:
        
        try:
            
            transport['Rate'] = int(input(prompt))
            break
        
        except ValueError:
            
            print('Enter a number!')
    
    return transport

