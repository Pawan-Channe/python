
print('\nCalculator\n')
num=int(input('Enter first number '))
num2=int(input('Enter second number '))
ope=input('Enter operator ')
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
    print('Enter a valid operator')

