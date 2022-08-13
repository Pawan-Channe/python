




question_li=["How many continents are there?","What is the capital of India?","NG mei kaun se course padhaya jaata hai?"]
option_li=[["1 Seven", "2 Nine", "3 Four", "4 Eight"],["1 Chandigarh", "2 Bhopal", "3 Chennai", "4 Delhi"],["1 Software Engineering", "2 Counseling", "3 Tourism", "4 Agriculture"]]
for number in range(len(question_li)):
    question=question_li[number]
    print(question)
    for option in option_li[number]:
        nested=option_li[number]
        print(option)
    print()
    answer=input("enter your answer:")
    if question_li.index(question)==0 and (answer=='1' or answer=='Seven'):
        print('"Congratulation" you won:10,000\n')
    elif question_li.index(question)==1 and (answer=='4' or answer=='Delhi'):
        print('"Congratulation" you won:20,000\n')
    elif question_li.index(question)==2 and (answer=='1' or answer=='Software Engineering'):
        print('"Congratulations" you won:30,000\n')
    else:
        print('"Opps wrong answer"\n')
        ask=input('Don you want lifeline:y/n:')
        if ask=='y' or ask=='yes':
            print('50-50')
            print(nested[0])
            print(nested[3],)
            choose=input('choose an option:')
            if question_li.index(question)==0 and (choose=='1' or choose=='Seven'):
                print('"Congratulation" you won:10,000\n')
            elif question_li.index(question)==1 and (choose=='4' or choose=='Delhi'):
                print('"Congratulation" you won:20,000\n')
            elif question_li.index(question)==2 and (choose=='1' or choose=='Software Engineering'):
                print('"Congratulations" you won:30,000\n')
            else:
                print('Your wrong')
                print('Your out of KBC')
                print('Bye..!')
                break
        else:
            print('Your out with No money..!')  
            break













