def cell_coord(cellname): #  преобразовывает имя ячейки АA32 в координаты
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    table = str.maketrans('0123456789','0123456789',  alphabet)
    digits = cellname.translate(table)
    table = str.maketrans(alphabet, alphabet,'0123456789')
    chars = cellname.upper().translate(table)
    #print(chars, digits)
    if not (chars.isalpha() and digits.isnumeric()): 
        raise ValueError( "invalid cellname not 'A1' format")
    nrow = int(digits)-1
    if len(chars) == 1:
        ncol = alphabet.index(chars)
    elif len(chars) == 2:
        ncol = (1+alphabet.index(chars[0]))*len(alphabet)+alphabet.index(chars[1])
    else:
        raise ValueError( "invalid cellname not 'A1' or 'AA123' format")
    return  (nrow,ncol)#     

def col_coord(colname): #  преобразовывает имя ячейки А32 в координаты
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    if len(colname) == 1:
        ncol = alphabet.index(colname)
    elif len(colname) == 2:
        ncol = (1+alphabet.index(colname[0]))*len(alphabet)+alphabet.index(colname[1])
    else:
        raise ValueError( "invalid cellname not 'A' or 'AA' format")
    return  ncol
