
new={}
while True:
    print('\npress 1 for create\npress 2 for read\npress 3 for update\npress 4 for delete\npress 5 for exit\n')
    choice=int(input("Enter your choice : "))
    if choice == 1:
        user_id = int(input('Enter your id : '))
        if user_id not in new:   
            name = input("enter your name : ")
            email=input("enter your email : ")
            if ('@' in email) and ('.com' in email or '.org' in email):
                number=input("enter your number : ")
                if len(number)==10:
                    new[user_id]={'name':name,'email':email,'number':int(number)}
                    print('Your data created successfully...')  
                else:
                    print("\nInvalid phone number")
                    print("Please try again..!\n")
            else:
                print("\nInvalid email")
                print("Please try again..!\n")
        else:
            print("\nYour id already exist")
            print("Please try again..!\n")
    elif choice == 2:
        main_key=int(input("Enter your id : "))
        if main_key in new:
            print(new[main_key])
        else:
            print('\nYour id does not exist')
    elif choice == 3:
        update=int(input("Enter your id : "))
        if update in new:
            print('\npress 1 for name\npress 2 for email\npress 3 for number\n')
            upgrade=int(input("Enter your choice : "))
            if upgrade == 1:
                label=input("enter your new name : ")
                new[update]['name']=label
                print("Your name updated successfully...")
            elif upgrade == 2:
                mail=input("enter your new email : ")
                if ('@' in mail) and ('.com' in mail or '.org' in mail):
                    new[update]['email']=mail
                    print("Your email updated successfully...")
                else:
                    print("\nInvalid email")
                    print("Please try again..!\n")
            elif upgrade == 3:
                numeral=input("enter your new number : ")
                if len(numeral)==10:
                    new[update]['number']=int(numeral)
                    print("Your number updated successfully...")
                else:
                    print("\nInvalid phone number")
                    print("Please try again..!\n")
        else:
            print('\nYour id does not exist')
    elif choice == 4:
        delete_user_id = int(input('Enter your id : '))
        if delete_user_id in new:
            sure = input('Are you sure to delete your account : yes or no : ')
            if sure == 'yes':
                del new[delete_user_id]
                print('Your account deleted successfully...')
        else:
            print('\nid does not exist')
    elif choice == 5:
        print("Your out of CRUD..!")
        break
    else:
        print("\nOops choice does not exist")
        print("Please try again..!\n")

        

