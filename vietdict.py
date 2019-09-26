
import xlrd

# Give the location of the file
loc = ("dictionary.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
# For row 0 and column 0
#print(sheet.cell_value(0, 0))
# Extracting number of rows
#print(sheet.nrows)

def random_word(minval,maxval,selection):

    import random
    if maxval>sheet.nrows:
        i = random.randint(minval-1,sheet.nrows-1)
    else:
        i = random.randint(minval-1,maxval-1)
    eng = str(sheet.cell_value(i,1))
    viet = str(sheet.cell_value(i,0))
    eng2 = 'EmptyCell'
    viet2 = 'EmptyCell'
    if sheet.cell_type(i,3)!=xlrd.XL_CELL_EMPTY:#Checking to see if the Viet word has multiple meanings
        eng2 = str(sheet.cell_value(i,3))

    if sheet.cell_type(i,2)!=xlrd.XL_CELL_EMPTY:#Checking to see if the English word has multiple meanings
        viet2 = str(sheet.cell_value(i,2))



    if selection=='1':#Translate Enlish to Vietnamese
        print('\nTranslate: ',eng)
        x = str(input())
        if x == viet or x == viet2:
            print('Correct: ', viet)
            if viet2!='EmptyCell':
                print('or: ',viet2)
            return random_word(minval,maxval,selection)
        elif x == 'stop':
            print('Quitting')
            return
        else:
            print('Incorrect. Try again:')
            y = str(input())
            if y == viet or y == viet2:
                print('Correct: ', viet)
                if viet2!='EmptyCell':
                    print('or: ',viet2)


            elif y == 'stop':
                print('Quitting')
                return
            else:
                print('Incorrect.\nCorrect answer is:', viet )
                if viet2!='EmptyCell':
                    print('or: ',viet2)
                return random_word(minval,maxval,selection)
        return random_word(minval,maxval,selection)
    elif selection=='2':#Translate Vietnamese to English
        print('\nTranslate: ',viet)
        x = str(input())
        if x == eng or x == eng2:
            print('Correct: ', eng)
            if eng2!='EmptyCell':
                print('or: ',eng2)
            return random_word(minval,maxval,selection)
        elif x == 'stop':
            print('Quitting')
            return
        else:
            print('Incorrect. Try again:')
            y = str(input())
            if y == eng or y == eng2:
                print('Correct: ', eng)
                if eng2!='EmptyCell':
                    print('or: ',eng2)
            elif y == 'stop':
                print('Quitting')
                return
            else:
                print('Incorrect.\nCorrect answer is:', eng )
                if eng2!='EmptyCell':
                    print('or: ',eng2)
                return random_word(minval,maxval,selection)
        return random_word(minval,maxval,selection)






selection=str(input('Choose: \n1. Translate English to Vietnamese\n2. Translate Vietnamese to English\n'))
if selection!='1' and selection!='2':
    print('No you have to enter 1 or 2')
    exit()
minval = int(input('Enter the range of numbers you want to practice.\nMin:'))
maxval = int(input('Max:'))
random_word(minval,maxval,selection)
