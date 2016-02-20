#import xlsxwriter
from xlsxwriter.utility import xl_range, xl_cell_to_rowcol

from .functions import draw_labels, draw_predm_table_header, draw_table_body #, simple_table
from .kalendar import  draw_kalendar_plan
#from .kalendar import draw_kalendar_plan

##    TP_INFO = {
##        'decan_pos'  :"B166" , 
##        'zavkav_pos' : "B169",
##        'decan_name'  :"C166" , 
##        'zavkav_name'   : "C169",
##        'prorector_pos' :"N166" , 
##        'HMB_pos'       : "N169",
##        'prorector_name' :"Y166" , 
##        'HMB_name'       : "Y169",
##        'year'
##      'spec_code'
##      'spec_name'
##      'kalendar'
##    }

def write_rnp(worksheet, rnp_data, STYLES):
    
    worksheet.set_landscape()
    worksheet.set_paper(9)
    #set_margins([left=0.7,] right=0.7,] top=0.75,] bottom=0.75]]])
    worksheet.set_margins(left=0.315, right=0.315, top=0.348, bottom=0.354)
    worksheet.center_horizontally()
    worksheet.set_print_scale(78)

    worksheet.set_column('A:BE', 2.4)
    worksheet.set_column('B:D', 2.1)
    worksheet.set_column('Q:Z', 2.1)
    worksheet.set_column('AM:AN', 2.6)
    worksheet.set_column('BA:BB', 2.6)

    worksheet.set_row(0, 14)
    #worksheet.set_row(start_row+3, 64)
    for n_row in range(12,175):
        worksheet.set_row(n_row, 12)
    
#rnp_data = (TP_INFO, get_course_TP(1,predmets),1)
    TP_INFO,PREDMETS, COURSE = rnp_data
    YEAR = str(TP_INFO['year'])

    UPHEADER_DATA = ( ("C1:L1","ЗАТВЕРДЖУЮ",'t12:cv:b' ),
                      ("C3","Проректор", 't12:lv:b'),
                      ("H3",TP_INFO['prorector_name'], 't12:lv:b'),
                      ("C5",'"______"_______________%s р.' % YEAR, 't12:lv:b'),
                      ("S1:AN1","Одеський національний політехнічний університет", 't12:cv:b'),
                      ("V3:AK3","Робочий навчальний план", 't12:cv:b'),
                      ("U5","Напряму", 't12:lv:_'),
                      ("Y5:AR5",'{} \u2014 {}'.format(TP_INFO['spec_code'], TP_INFO['spec_name']), 't12:lv:ib'),
                      ("U7","Курс", 't12:lv:_'),
                      ("Y7",COURSE, 't12:lv:bi'),
                    )

    draw_labels(worksheet,UPHEADER_DATA, STYLES)

    #draw_kalendar_plan(worksheet, start_pos, DATA, STYLES):
    WEEK_DATA = TP_INFO['kalendar'][COURSE-1]
    draw_kalendar_plan(worksheet, "B9", [(COURSE, WEEK_DATA)], STYLES)

    sem_header_data = [ ("2:3", "всього", 't8:cv:wr:1'),
                         [('5:1', "Обсяг ауд занять", 't8:cv:w:1'),
                              ('2:2', "всього", 't8:cv:wr:1'),
                              [    ('3:1', "в тому числі", 't8:cv:w:1'),
                                   ('1:1', "лекції", 't8:cv:wr:1'),
                                   ('1:1', "практичні", 't8:cv:wr:1'),
                                   ('1:1', "лабораторні", 't8:cv:wr:1'),
                               ]
                          ],
                         ('2:3', 'самостійна робота', 't8:cv:wr:1'),
                         ('1:3', "курсові роботи, проекти", 't8:cv:wr:1'),
                         ('1:3', "розрахункові роботи", 't8:cv:wr:1'),
                         ('1:3', "К-ть кредитів", 't8:cv:wr:1'),
                         [('2:2', "форми контролю", 't8:cv:w:1'),
                              ('1:1', "екзамен", 't8:cv:wr:1'),
                              ('1:1', "залік", 't8:cv:wr:1'),
                          ]
                    ]

    header_data = (("1:4", "№ з/п", 't8:cv:wr:1'),
                   ('3:4', "Шифр за ОПП", 't8:cv:wr:1'),
                   ('12:4', 'Назва навчальних дисциплін', 't8:cv:w:1'),
                   ('2:4', "К-ть кредитів", 't8:cv:wr:1'),
                        [   ('8:1', "Кількість годин", 't8:cv:w:1'),
                            ('2:3', "за навчальним планом",  't8:cv:wr:1'),
                            ('2:3', "фактично виділено",     't8:cv:wr:1'),
                            ('2:3', "прочитано в минулому році", 't8:cv:wr:1'),
                            ('2:3',"на поточний навчальний рік", 't8:cv:wr:1'),
                        ],
                   [('14:1', "1 семестр  18 навчальних тижнів", 't8:cv:w:1')]+sem_header_data,
                   [('14:1', "2 семестр  18 навчальних тижнів", 't8:cv:w:1')]+list(sem_header_data),
                   ('3:4', "Кафедра", 't8:cv:w:1'),
                )

    
    cur_row = draw_predm_table_header(worksheet, 'A14', header_data, STYLES)

    DELTAS = 1,3,12, 2,2,2,2,2,   2,2,1,1,1,2,1,1,1,1,1,   2,2,1,1,1,2,1,1,1,1,1, 3
    table_styles = {'def':"t8:cv:_:1", 2:"t8:lv:w:1", 0:"t7:cv:_:1"}
    attrib = (table_styles, DELTAS) 
    

    cur_row = draw_table_body(worksheet, (cur_row+1, 0), PREDMETS, attrib, STYLES)


    def get_itogo(PREDMETS):
        def get_column(n):
            res = []
            for item in PREDMETS:
                res.append(item[n])
            return res
        def sum_column(col):
            ss=0
            for s in col:
                if not s:
                    s = 0
                elif type(s) == str:
                    s = 1
                ss +=s
            return ss
            
        out = ['','', 'Разом:',]
        for n in range (3, len(PREDMETS[0])-1):
            cur_col = get_column(n)
            cur =     sum_column(cur_col)
            out.append(cur)
        out.append('')
        return out

    foot = get_itogo(PREDMETS)
    print(foot)
    table_styles = {'def':"t9:cv:b:1", 2:"t9:rv:b:1", 0:"t7:cv:_:1"}
    attrib = (table_styles, DELTAS) 
    draw_table_body(worksheet, (cur_row, 0), [foot], attrib, STYLES)
   
                      
                
                      
                      
                      
                      
                      
                      
                      







    
