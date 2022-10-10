import json,os
print('\n\033[1m',"WELCOME TO CRUD OPERATION",'\033[0m')
while True:
    try:
        print()
        print("\n press 1 for create \n press 2 for read \n press 3 for update \n press 4 for delete \n press 5 for exit",'\033[0m')
        print()
        def create():
            if os.path.exists("crud.json"):
                a = input("Enter number: ")
                with open("crud.json","r") as obj1:
                    if obj1.read() == "":
                        d1 ={}
                    else:
                        obj1.seek(0)
                        d1 = json.load(obj1)
                    d1[a]=({"name":input("Enter your name: "),
                    "email":input("Enter your email: ")})
                    with open("crud.json","w") as obj1:
                        json.dump(d1,obj1,indent=4)
                        print()
                        print('\033[35m'," Your info created successfully ",'\033[0m')
            else:
                with open('crud.json','w'):
                    create()
        def read():
            b = (input("Enter your number: "))
            with open("crud.json","r") as red:
                lo = json.load(red)
                if b in lo:
                    print()
                    print(lo[b])
                else:
                    print()
                    print('\033[1m','\033[91m',"Number does not exist: ",'\033[0m')
                    read()
        def update():
            v = (input('Which number of data do you want to update: '))
            with open("crud.json","r") as obj3:
                data = json.load(obj3)
                if v in data:
                    a = {"name":input("Enter your name: "),"email":input("Enter your email: ")}
                    data[v] = a
                    with open("crud.json","w") as obj4:
                        json.dump(data,obj4,indent=4)
                        print()
                        print('\033[37m',"Updated Successfully.......",'\033[0m')
                else:
                    print()
                    print('\033[37m',"Your number does not exist",'\033[0m')
                    update()
        def delete():
            with open("crud.json","r") as data:
                d1 = json.load(data)
            m = (input("Which number of data do you want to delete: "))
            if m in d1:
                d1.pop(m)
                with open("crud.json","w") as obj5:
                    json.dump(d1,obj5,indent=4)
                    print()
                    print('\033[36m',"Successfully Deleted",'\033[0m')
            else:
                print()
                print('\033[36m',"Number does not exist: ",'\033[0m')
                delete()
        choice = int(input("Enter your choice: "))
        if choice == 1:
            create()
        elif choice == 2:
            read()
        elif choice == 3:
            update()
        elif choice == 4:
            delete()
        else:
            print("You are out of CRUD..!")
            break
    except:
        print()
        print('\033[37m',"Your mobail number does not exist",'\033[0m')