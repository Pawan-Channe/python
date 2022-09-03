
print("\nDIGITAL_FM\n")
print('Do you like to eat?\npress 1 for Yes\npress 2 for No\n')
a=input('Enter:')
if a=='1' or a=='Yes':
    print('\nWelcome to hotel: RED APPLE..!')
    print('1.breakfast\n2.lunch\n3.dinner')
    b=input('\nplease choose a meal type:')
    if b=='1' or b=='breakfast':
        print("1.tea\n2.poha\n3.sandwitch")
        c=input("\nplease select a breakfast:")
        if c=='1' or c=='tea':
            print('\nMenu and Rate:')
            print('1.green tea : 20rs\n2.coffe : 30rs\n3.black tea : 10rs\n4.milk tea : 25rs')
            p=input('\ngiven above which tea you want?:')
            if p=='1' or p=='green tea':
                print('1.full\n2.half')
                q=input('choose to buy:')
                if q=='1' or q=='full':
                    print('Total Amount: 20rs')
                else:
                    print('Total Amount: 10rs')
            elif p=='2' or p=='coffe':
                print('1.full\n2.half')
                r=input('choose to buy:')
                if r=='1' or r=='full':
                    print('Total Amount: 30rs')
                else:
                    print('Total Amount: 15rs')
            elif p=='3' or p=='black tea':
                print('1.full\n2.half')
                s=input('choose to buy:')
                if s=='1' or s=='full':
                    print('Total Amount: 10rs')
                else:
                    print('Total Amount: 5rs')
            elif p=='4' or p=='milk tea':
                print('1.full\n2.half')
                t=input('choose to buy:')
                if t=='1' or t=='full':
                    print('Total Amount: 25rs')
                else:
                    print('Total Amount: 12.5rs')
            else:
                print("Sorry this tea is not exist..!")
        elif c=='2' or c=='poha':
            print('Menu and Rate:')
            print('1.aalu poha : 30rs\n2.fresh poha : 40rs\n3.egg poha : 50rs')
            w=input('\ngiven above which poha you want?:')
            if w=='1' or w=='aalu poha':
                print('1.full\n2.half')
                input('choose to buy:')
                u=input('choose to buy:')
                if u=='1' or u=='full':
                    print('Total Amount: 30rs')
                else:
                    print('Total Amount: 15rs')
            elif w=='2' or w=='fresh poha':
                print('1.full\n2.half')
                v=input('choose to buy:')
                if v=='1' or v=='full':
                    print('Total Amount: 40rs')
                else:
                    print('Total Amount: 20rs')
            elif w=='3' or w=='egg poha':
                print('1.full\n2.half')
                x=input('choose to buy:')
                if x=='1' or x=='full':
                    print('Total Amount: 50rs') 
                else:
                    print('Total Amount: 25rs')
            else:
                print("Sorry this poha is not exist..!")    
        elif c=='3' or c=='sandwitch':
            print('Menu and Rate:')
            print('1.non-veg sandwitch : 60rs\n2.veg sandwitch : 30')
            y=input('given above which sandwitch you want?')
            if y=='1' or y=='non-veg sandwitch':
                print('1.full\n2.half')
                z=input('choose to buy:')
                if z=='1' or z=='full':
                    print('Total Amount: 60rs')
                else:
                    print('Total Amount: 30rs')
            elif y=='2' or y=='veg sandwitch':
                print('1.full\n2.half')
                a1=input('choose to buy:')
                if a1=='1' or 'full':
                    print('Total Amount: 30rs')
                else:
                    print('Total Amount: 15rs')
            else:
                print("Sorry this sandwich is not exist..!")
        else:
            print("Oops something went wrong..!")
    elif b=='2' or b=='lunch':
        print("1.veg chawal\n2.non-veg chawal")
        d=input("\nplease select a lunch:")
        if d=='1' or d=='veg chawal':
            print('Menu and Rate:')
            print('1.daal chawal : 40rs\n2.paneer chawal : 70rs')
            a2=input('given above which veg-chawal you want?:')
            if a2=='1' or a2=='daal chawal':
                print('1.full\n2.half')
                a3=input('choose to buy:')
                if a3=='1' or a3=='full':
                    print('Total Amount: 40rs')
                else:
                    print('Total Amount: 20rs')
            elif a2=='2' or a2=='paneer chawal':
                print('1.full\n2.half')
                a4=input('choose to buy:')
                if a4=='1' or a4=='full':
                    print('Total Amount: 70rs')
                else:
                    print('Total Amount: 35rs')
            else:
                print("Sorry this chawal is not exist..!")
        elif d=='2' or d=='non-veg chawal':
            print('Menu and Rate:')
            print('1.chicken chawal : 80rs\n2.egg chawal : 60rs')
            a4=input('given above which non-veg chawal you want?:')
            if a4=='1' or a4=='chicken chawal':
                print('1.full\n2.half')
                a5=input('choose to buy:')
                if a5=='1' or a5=='full':
                    print('Total Amount: 80rs')
                else:
                    print('Total Amount: 40rs')
            elif a4=='2' or a4=='egg chawal':
                print('1.full\n2.half')
                a6=input('choose to buy:')
                if a6=='1' or a6=='full':
                    print('Total Amount: 60rs')
                else:
                    print('Total Amount: 30rs')
            else:
                print("Sorry this tea is not exist..!")
        else:
            print("Oops something went wrong..!")
    elif b=='3' or b=='dinner':
        print("1.veg\n2.non-veg")
        e=input("\nplease select a dinner:")
        if e=='1' or e=='veg':
            print('Menu and Rate:')
            print('1.flower roti : 60rs\n2.paneer roti : 60rs\n3.brinjal roti : 40rs\n4.potato roti : 40rs')
            a7=input('given above which veg-roti you want?')
            if a7=='1' or a7=='flower roti' or a7=='2' or a7=='paneer roti':
                print('1.full\n2.half')
                a8=input('choose to buy:')
                if a8=='1' or a8=='full':
                    print('Total Amount: 60rs')
                else:
                    print('Total Amount: 30rs')
            elif a7=='3' or a7=='brinjal roti' or a7=='4' or a7=='potato roti':
                print('1.full\n2.half')
                a10=input('choose to buy:')
                if a10=='1' or a10=='full':
                    print('Total Amount: 40rs')
                else:
                    print('Total Amount: 20rs')
            else:
                print("Sorry this veg-roti is not exist..!")
        elif e=='2' or e=='non-veg':
            print('Menu and Rate:')
            print('1.chicken roti : 100rs\n2.eggs roti : 50rs')
            a11=input('given above which non-veg roti you want?:')
            if a11=='1' or a11=='chicken roti':
                print('1.full\n2.half')
                a12=input('choose to buy:')
                if a12=='1' or a12=='full':
                    print('Total Amount: 100rs')
                else:
                    print('Total Amount: 50rs')
            elif a11=='2' or a11=='eggs roti':
                print('1.full\n2.half')
                a15=input('choose to buy:')
                if a15=='1' or a15=='full':
                    print('Total Amount: 50rs')
                else:
                    print('Total Amount: 25rs')
            else:
                print("Sorry this non-veg roti is not exist..!")
        else:
            print("Oops something went wrong..!")
    else:
        print("Oops something went wrong..!")
elif a=='2' or a=='No':
    print("\nGood bye..!")
else:
    print("\nOops You entered something else:")


