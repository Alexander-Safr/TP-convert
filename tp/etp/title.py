from .functions import draw_labels , simple_table 
from .kalendar import draw_kalendar_plan


def title_header(worksheet, TP_ATTRIB, TP_TITLE_TABLES, STYLES):

    #TP_ATTRIB, TP_TITLE_TABLES = METADATA

    YEAR = TP_ATTRIB['year'][1]
    TP_ATTRIB.pop('year',None)

    TIT_STYLE_MAP = {
'level': "t11:r:bi", "domain_code" : "t11:l:bi",# 'domain_name' :"U8:AL8", "qualification" : "AR8:BA8",
##    'trend_code': "F9:I9", 'trend_name': "K9:AL9", 'trend': "AN9:BA9",
##    'spec_code': "F10:J10", 'spec_name': "L10:AL10", 'term': "AS10:BA10",
##    
##    "specialization" : "F11:AL11", "base" : "AQ11:BA11",
##    "teach_form" : "O12:AL12",
   }

##    def get_style(stl = 't11:l:_'):
##        return get_style_(STYLES, stl)

    def get_metadata(TP_ATTRIB,TIT_STYLE_MAP, style = 't11:c:_'):
        out = []
        for key in TP_ATTRIB:
            cur_style = TIT_STYLE_MAP.get(key, style)
            out.append((TP_ATTRIB[key][0],TP_ATTRIB[key][1], cur_style))
        return out
            

    worksheet.set_landscape()
    worksheet.set_paper(9)
    #set_margins([left=0.7,] right=0.7,] top=0.75,] bottom=0.75]]])
    worksheet.set_margins(left=0.315, right=0.315, top=0.348, bottom=0.354)
    worksheet.center_horizontally()
    worksheet.set_print_scale(83)

    worksheet.set_column('A:BE', 2.4)

    for n_row in range(0,40):
        worksheet.set_row(n_row, 13)
    worksheet.set_row(29, 74)
    

    LABELS_DATA = ( ("D1","ЗАТВЕРДЖУЮ",'t11:c:b' ) , ("P2:AK2", "Міністерство освіти і науки України",'t11:c:b'),
             ("A3","Ректор                          Г.О. Оборський",'t11:c:b'),
             ("P3:AK3", 'Одеський національний політехнічний університет','t11:c:b'),
             ("A5:M5", '"______"_______________ {} р.'.format(YEAR),'t11:c:b'),
             ("P6:AK6", 'НАВЧАЛЬНИЙ    ПЛАН','t11:c:b'),
             ("A8", "підготовки"), ("L8","з галузі знань"), ("AN8",'Кваліфікація'),
             ("A9", "за напрямом"),
             ("A10", "спеціальністю"), ("AN10",'Строк навчання'),
             ("A11", "спеціалізацією"), ("AN11",'на основі'),
             ("I12", "Форма навчання"),
             ("A14:BA14", "І. ГРАФІК НАВЧАЛЬНОГО ПРОЦЕСУ","t11:c:b"),
             ("B28:R28", "ІІ. ЗВЕДЕНІ ДАНІ ПРО БЮДЖЕТ ЧАСУ, тижні","t11:c:b"),
             ("V28:AF28", "ІІІ. ПРАКТИКА","t11:c:b"),
             ("AJ28:AZ28", "IV. ДЕРЖАВНА АТЕСТАЦІЯ","t11:c:b"),
             
            )


    draw_labels(worksheet, LABELS_DATA, STYLES)

    FIELD_DATA = get_metadata( TP_ATTRIB,TIT_STYLE_MAP, style = 't11:c:bi')
    draw_labels(worksheet,FIELD_DATA, STYLES)

    kaldata = [(c+1,TP_TITLE_TABLES['kalendar'][c]) for c in range(6)]
    #kaldata = (1, 'T'*52),(2, 'T'*52),(3, 'T'*52),(4, 'T'*40),(5, ''),(6, '')
    draw_kalendar_plan(worksheet, "A16", kaldata, STYLES)
    
    #print(*TP_TITLE_TABLES['budget'], sep = '\n')
    
    table_data = {'start' : "B30",
                  'header': (('Курс',"t10:cv:r:1" ),'Теоретичне навчання','Екзаменаційна сесія',
                             'Практика','Державна атестація','Виконання дипломного проекту (роботи)',
                             'Канікули','Разом'),
                  'header_sizes': (2,2,2,2,2,3,2,2),
                  'header_style': "t9:cv:rw:1",
                  'body_style'  : "t9:cv:_:1",
                  'data'        : TP_TITLE_TABLES['budget'] 
                  }
    simple_table(worksheet, table_data, STYLES)

    table_data = {'start' : "V30",
        'header': (('Назва практики',"t9:cv:_:1"),'Семестр','Тижні',
                             ),
                  'header_sizes': (7,2,2),
                  'header_style': "t9:cv:rw:1",
                  'body_style'  : "t9:cv:_:1",
                  'data'        : TP_TITLE_TABLES['practic'] 
                  }
    simple_table(worksheet,  table_data, STYLES)

    table_data = {'start' : "AJ30",
        'header': ('Назва навчальної дисципліни',
                             'Форма державної атестації (екзамен, дипломний проект (робота))',
                             ('Семестр',"t9:cv:r:1" ),
                             ),
                  'header_sizes': (7,8,2),
                  'header_style': "t9:cv:w:1",
                  'body_style'  : "t9:cv:_:1",
                  'data'        : TP_TITLE_TABLES['attest'] 
                  }
    simple_table(worksheet,table_data, STYLES)


        
        


    
