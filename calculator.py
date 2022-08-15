

print('Calculator')
num=int(input('enter first number '))
num2=int(input('enter second number '))
ope=input('enter operator ')
if ope=='+':
    print(num+num2)
    print('Addition')
elif ope=='-':
    print(num-num2)
    print('Subtraction')
elif ope=='*':
    print(num*num2)
    print('Multiplication')
elif ope=='/':
    print(num/num2)
    print('Division')
elif ope=='//':
    print(num//num2)
    print('Floor division')
elif ope=='%':
    print(num%num2)
    print('Modulus')
else:
    print('enter a valid operator')
    
    
