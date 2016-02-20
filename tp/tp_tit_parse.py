import xlrd
import pickle
from xlsxwriter.utility import xl_range, xl_cell_to_rowcol as coord
# xl_cell_to_rowcol()
# (row, col) = xl_cell_to_rowcol('B1') # (0, 1)
# cell_range = xl_range(0, 0, 3, 4) # A1:E4

#SHEET_TITLE_NAME = "ТитулНП"
#FILENAME = 'TP2.xls'


#wb = xlrd.open_workbook(FILENAME,formatting_info=False)
#tp_title = wb.sheet_by_name(SHEET_TITLE_NAME)
def tit_parse(tp_title):
    TITUL_MAP = {
        'year': "L5",
        'level': "F8:J8", "domain_code" : "Q8:S8", 'domain_name' :"U8:AL8", "qualification" : "AR8:BA8",
        'trend_code': "F9:I9", 'trend_name': "K9:AL9", 'trend': "AN9:BA9",
        'spec_code': "F10:J10", 'spec_name': "L10:AL10", 'term': "AS10:BA10",
        
        "specialization" : "F11:AL11", "base" : "AQ11:BA11",
        "teach_form" : "O12:AL12",
     }

    TITUL_BLOCK_MAP = {
        "budget" : ("B31:R37",   (2,2,2,2,2,3,2,2)), 
        "practic": ("V31:AF33",  (7,2,2)),
        "attest" : ("AJ31:AZ32", (7,8,2)),
        'kalendar' :("B18:BA23", (1,)*52)
        }

    def extract_tit_data(tp_title, TITUL_MAP, TITUL_BLOCK_MAP):
        def part_cellname(string):
            return string.split(':')[0]



        def unmerge(l, mask):

            def take_one(lst):
                if not lst:
                    return ''
                for i in lst:
                    if i:
                        if type(i) == float:
                            return int(i)
                        else:
                            return i
                return ''
            
            data = l.copy()
            #print(data)
            out = []
            for m in mask:
                out.append(take_one([data.pop(0) for x in range(m)]))
            #print(out)
            return out

        def extract_block(crd, mask):
            start, end = crd.split(':')
            sr,sc = coord(start)
            er,ec = coord(end)
            out = []
            #row_values(rowx, start_colx=0, end_colx=None)
            for row in range(sr,er+1):
                out.append(unmerge(tp_title.row_values(row, start_colx=sc, end_colx=ec+1), mask))
            return out
            
        TP_ATTRIB       = dict(TITUL_MAP)
        TP_TITLE_TABLES = dict(TITUL_BLOCK_MAP)
        
        for key in TITUL_MAP:
            key = part_cellname(key)
            TP_ATTRIB[key] = (TITUL_MAP[key],tp_title.cell(*coord(TITUL_MAP[key])).value)
            
        TP_ATTRIB['year'] = list(TP_ATTRIB['year'])
        TP_ATTRIB['year'][1] = 2000+int(TP_ATTRIB['year'][1] %100)
        TP_ATTRIB['year'] = tuple(TP_ATTRIB['year'])
        

        for key in TITUL_BLOCK_MAP:
            crd, mask = TITUL_BLOCK_MAP[key]
            
            TP_TITLE_TABLES[key] = extract_block(crd, mask)
            #TP_TITLE_TABLES[key] = extract_block(crd)

        return TP_ATTRIB, TP_TITLE_TABLES

    TP_ATTRIB, TP_TITLE_TABLES = extract_tit_data(tp_title, TITUL_MAP, TITUL_BLOCK_MAP)

    #with open('title_data2.pickle', 'wb') as f:
    #    pickle.dump((TP_ATTRIB, TP_TITLE_TABLES), f)
    return TP_ATTRIB, TP_TITLE_TABLES 
    
