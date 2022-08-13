
a=input('Do you like to eat? 1.yes or 2.no? ')
if a=='1' or a=='yes':
    print('welcome to hotel "Red apple"')
    b=input('please choose a meal type "1.breakfast" or "2.lunch" or "3.dinner" ')
    if b=='1' or b=='breakfast':
        c=input("""1.tea\n2.poha\n3.sandwitch\n""")
        if c=='1' or c=='tea':
            print(' menu and rate:')
            print('1.green tea = 20rs\n2.coffe = 30rs\n3.black tea = 10rs\n4.milk tea = 25rs')
            p=input('given above which tea you want? ')
            if p=='1' or p=='green tea':
                q=input('1.full or 2.half ')
                if q=='1' or q=='full':
                    print('20rs')
                else:
                    print('10rs')
            elif p=='2' or p=='coffe':
                r=input('1.full or 2. half ')
                if r=='1' or r=='full':
                    print('30rs')
                else:
                    print('15rs')
            elif p=='3' or p=='black tea':
                s=input('1.full or 2.half ')
                if s=='1' or s=='full':
                    print('10rs')
                else:
                    print('5rs')
            elif p=='4' or p=='milk tea':
                t=input('1.full or 2.half ')
                if t=='1' or t=='full':
                    print('25rs')
                else:
                    print('12.5rs')
        elif c=='2' or c=='poha':
            print('menu and rate:')
            print('1.aalu poha = 30rs\n2.fresh poha = 40rs\n3.egg poha = 50rs')
            w=input('given above which poha you want?')
            if w=='1' or w=='aalu poha':
                u=input('1.full or 2.half ')
                if u=='1' or u=='full':
                    print('30rs')
                else:
                    print('15rs')
            elif w=='2' or w=='fresh poha':
                v=input('1.full or 2.half ')
                if v=='1' or v=='full':
                    print('40rs')
                else:
                    print('20rs')
            elif w=='3' or w=='egg poha':
                x=input('1.full or 2.half ')
                if x=='1' or x=='full':
                    print('50rs') 
                else:
                    print('25rs')       
        elif c=='3' or c=='sandwitch':
            print('menu and rate:')
            print('1.non-veg sandwitch = 60rs\n2.veg sandwitch = 30')
            y=input('given above which sandwitch you want?')
            if y=='1' or y=='non-veg sandwitch':
                z=input('1.full or 2.half ')
                if z=='1' or z=='full':
                    print('60rs')
                else:
                    print('30rs')
            elif y=='2' or y=='veg sandwitch':
                a1=input('1.full or 2.half ')
                if a1=='1' or 'full':
                    print('30rs')
                else:
                    print('15rs')
            
    elif b=='2' or b=='lunch':
        d=input("""1.veg chawal\n2.non-veg chawal""")
        if d=='1' or d=='veg chawal':
            print('menu and rate:')
            print('1.daal chawal = 40rs\n2.paneer chawal = 70rs')
            a2=input('given above which veg-chawal you want?')
            if a2=='1' or a2=='daal chawal':
                a3=input('1.full or 2.half ')
                if a3=='1' or a3=='full':
                    print('40rs')
                else:
                    print('20rs')
            elif a2=='2' or a2=='paneer chawal':
                a4=input('1.full or 2.half ')
                if a4=='1' or a4=='full':
                    print('70rs')
                else:
                    print('35rs')
        elif d=='2' or d=='non-veg chawal':
            print('menu and rate:')
            print('1.chicken chawal = 80rs\n2.egg chawal = 60rs')
            a4=input('given above which non-veg chawal you want?')
            if a4=='1' or a4=='chicken chawal':
                a5=input('1.full or 2.half ')
                if a5=='1' or a5=='full':
                    print('80rs')
                else:
                    print('40rs')
            elif a4=='2' or a4=='egg chawal':
                a6=input('1.full or 2.half ')
                if a6=='1' or a6=='full':
                    print('60rs')
                else:
                    print('30rs')
    elif b=='3' or b=='dinner':
        e=input("""1.veg\n2.non-veg """)
        if e=='1' or e=='veg':
            print('menu and rate:')
            print('1.flower roti = 60rs\n2.paneer roti = 60rs\n3.brinjal roti = 40rs\n4.potato roti = 40rs')
            a7=input('given above which veg-roti you want?')
            if a7=='1' or a7=='flower roti' or a7=='2' or a7=='paneer roti':
                a8=input('1.full or 2.half ')
                if a8=='1' or a8=='full':
                    print('60rs')
                else:
                    print('30rs')
            elif a7=='3' or a7=='brinjal roti' or a7=='4' or a7=='potato roti':
                a10=input('1.full or 2.half ')
                if a10=='1' or a10=='full':
                    print('40rs')
                else:
                    print('20rs')
        elif e=='2' or e=='non-veg':
            print('menu and rate:')
            print('1.chicken roti = 100rs\n2.eggs roti = 50rs')
            a11=input('given above which non-veg roti you want?')
            if a11=='1' or a11=='chicken roti':
                a12=input('1.full or 2.half ')
                if a12=='1' or a12=='full':
                    print('100rs')
                else:
                    print('50rs')
            elif a11=='2' or a11=='eggs roti':
                a15=input('1.full or 2.half ')
                if a15=='1' or a15=='full':
                    print('50rs')
                else:
                    print('25rs')
else:
    print("Good bye..!")






