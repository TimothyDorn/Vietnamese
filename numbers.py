import xlrd
loc = ("numbers.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)


def smallNum(i):
    if i<=10:
        num = sheet.cell_value(i,1)
        return num
    elif i==15:
        num = 'mười lăm'
        return num
    elif i<=19:
        num = 'mười '+sheet.cell_value(i%10,1)
        return num
    elif i<=99:
        if i%10 == 0:
            num = sheet.cell_value(int(i/10),1) +' mươi'
        elif i%10==1:
            num = sheet.cell_value(int(i/10),1) +' mươi mốt'
        elif i%10==5:
            num = sheet.cell_value(int(i/10),1) +' mươi lăm'
        else:
            num = sheet.cell_value(int(i/10),1) +' mươi '+ sheet.cell_value((i%10),1)
        return num

def bigNum(i):
    if i<=999:
        if i%100 == 0:
            num = sheet.cell_value(int(i/100),1) + ' trăm'
        else:
            num = sheet.cell_value(int(i/100),1) + ' trăm '+ smallNum(i%100)
        return num


def random_VietToNum(minval,maxval):
    import random
    if maxval>999:
        i = random.randint(minval,999)
    else:
        i = random.randint(minval,maxval)

    if i<=99:
        num = smallNum(i)
    elif i<=999:
        num = bigNum(i)
    print('\nTranslate: ',num)
    x = str(input())
    if x == str(i):
        print('Correct: ', str(i))
        return random_VietToNum(minval,maxval)
    elif x == 'stop':
        print('Quitting')
        return
    else:
        print('Incorrect. Try again:')
        y = str(input())
        if y == str(i):
            print('Correct: ', str(i))
        elif y == 'stop':
            print('Quitting')
            return
        else:
            print('Incorrect.\nCorrect answer is:', str(i) )
            return random_VietToNum(minval,maxval)
    return random_VietToNum(minval,maxval)


def random_numToViet(minval,maxval):
    import random
    if maxval>999:
        i = random.randint(minval,999)
    else:
        i = random.randint(minval,maxval)

    if i<=99:
        num = smallNum(i)
    elif i<=999:
        num = bigNum(i)

    print('\nTranslate: ',i)
    x = str(input())
    if x == num:
        print('Correct: ', num)
        return random_numToViet(minval,maxval)
    elif x == 'stop':
        print('Quitting')
        return
    else:
        print('Incorrect. Try again:')
        y = str(input())
        if y == num:
            print('Correct: ', num)
        elif y == 'stop':
            print('Quitting')
            return
        else:
            print('Incorrect.\nCorrect answer is:', num )
            return random_numToViet(minval,maxval)
    return random_numToViet(minval,maxval)

x=str(input('Choose: \n1. Random number to Vietnamese\n2. Random Vietnamese to number\n'))

if x=='1':
    minval = int(input('Enter the range of numbers you want to practice.\nMin:'))
    maxval = int(input('Max:'))
    random_numToViet(minval,maxval)
elif x=='2':
    minval = int(input('Enter the range of numbers you want to practice.\nMin:'))
    maxval = int(input('Max:'))
    random_VietToNum(minval,maxval)
else:
    print('No you have to enter 1 or 2')
