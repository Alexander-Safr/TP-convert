from .etp_styles import get_style_
from .functions  import draw_text_rect, draw_many_text_rect
from xlsxwriter.utility import xl_cell_to_rowcol
#draw_text_rect вроде не надо


def draw_kalendar_plan(worksheet, start_pos, DATA, STYLES):

    def style(stl = 't10:c:_:1'):
        return get_style_(STYLES, stl)
    
    def draw_kalendar_plan_header():
        draw_many_text_rect(worksheet,(sy, sx+1),"4:1", MONTH, style(), (4,4,5,4,5,4,4, 4,5, 4,5,4))
        draw_many_text_rect(worksheet,(sy+1,sx+1, ),"1:1", range(1,53), style())
        worksheet.merge_range(sy, sx, sy+1, sx, "Курс", style('t10:c:r:1'))
    def draw_kalendar_plan_body(course,week):
        worksheet.write(sy+2+nrow,sx,course,style())
        draw_many_text_rect(worksheet, (sy+2+nrow, sx+1),"1:1", week, style())
        
    #sx, sy = dec_int(start_pos) if type(start_pos) == str else start_pos
    sy,sx = xl_cell_to_rowcol(start_pos)
    #style_kalendar, style_kalendar_bold, style(), style_legend = styles
    #COURSE,WEEK_DATA = DATA 
    MONTH = ('вересень','жовтень','листопад','грудень',
             'січень'  ,'лютий'  ,'березень','квітень',
             'травень' , 'червень','липень', 'серпень')
    LEGEND = """Позначення: Т - теоретичне навчання; С - екзаменаційна сесія;\
    П - практика; К - канікули; ДЕ - складання державного екзамену;\
    ДП - захист дипломного проекту (роботи), дп - виконання дипломного проекту (роботи)"""

    draw_kalendar_plan_header()
    nrow = 0
    for course, week in DATA:
        if len(week)<52:
            week = week + ' '*(52-len(week))
        draw_kalendar_plan_body(course, week)
        nrow+=1
    worksheet.write(sy+2+len(DATA),sx+1,LEGEND, style('t9:l:_'))
