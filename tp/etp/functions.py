from xlsxwriter.utility import xl_cell_to_rowcol
from .etp_styles import get_style_ 

def get_style(STYLES, stl = 't11:l:_'):
    return get_style_(STYLES, stl)

def dec_int(s):
    x,y = s.split(':')
    return int(x),int(y)

def draw_text_rect(worksheet, start_pos,dim, title, style):
    sy, sx = start_pos
    w, h   = dim
    if w == h == 1:
        worksheet.write(sy, sx, title,style)
    else:        
        worksheet.merge_range(sy, sx, sy+h-1, sx +w-1, title, style)
       
# old
def _draw_many_text_rect(worksheet, start_pos,dim, data, style, rdim = False):

    sx, sy  = dec_int(start_pos) if type(start_pos) == str else start_pos
    w, h    = dec_int(dim)       if type(dim) == str       else dim
    if not rdim:
        w = [w] * len(data)
    else:
        w = rdim
   
    for n,title in enumerate(data):
        draw_text_rect(worksheet, (sy,sx),(w[n],h), title, style)
        sx+=w[n]
# new
def draw_many_text_rect(worksheet, start_pos,dim, data, style, rdim = False):

    if type(start_pos) == str and ":" not in start_pos:
        sy,sx = xl_cell_to_rowcol(start_pos)
    else:
        sx, sy  = dec_int(start_pos) if type(start_pos) == str else start_pos[::-1]
    w, h    = dec_int(dim)       if type(dim) == str       else dim
    if not rdim:
        w = [w] * len(data)
    else:
        w = rdim
   
    for n,title in enumerate(data):
        draw_text_rect(worksheet, (sy,sx),(w[n],h), title, style)
        sx+=w[n]

# нарисовать несколько надписей, ДАТА (координаты, текст, стиль)
def draw_labels(worksheet,DATA, STYLES):
    for coord, label, *style in DATA:
        if style:
            stl = get_style(STYLES,style[0])
        else:
            stl = get_style(STYLES, 't11:l:_')
        if ":" in coord:
            worksheet.merge_range(coord, label, stl)
        else:
            worksheet.write(coord, label, stl)    
 

def simple_table(worksheet,  table_data, STYLES):
    start_pos = table_data['start']
    crow, ccol = srow, scol = xl_cell_to_rowcol(start_pos)
    assert len(table_data['header_sizes'])== len(table_data['header']), "неправильное наполннеие таблицы"

    def_hstyle =table_data['header_style']
    def_bstyle = get_style(STYLES, table_data['body_style'])
    hdata = table_data['header']
    
    for n in range(len(hdata)):
        if type(hdata[n])==str:
            h_stl,title  = def_hstyle, hdata[n]
        else:
            h_stl,title  = hdata[n][1], hdata[n][0]
        hstl = get_style(STYLES, h_stl)
       # print(h_stl, title)
        draw_text_rect(worksheet, (crow, ccol),(table_data['header_sizes'][n],1), title, hstl)
        ccol+= table_data['header_sizes'][n]

    crow, ccol = srow+1, scol
    #draw_many_text_rect(worksheet, start_pos,dim, data, style, rdim = False)
    for datarow in table_data['data']:
        datarow = [int(y) if (type(y)==str and y.isdigit()) else y for y in datarow]
        draw_many_text_rect(worksheet, (crow, ccol),"1:1", datarow, def_bstyle,
                        rdim = table_data['header_sizes'])
        crow+=1

        

def draw_predm_table_header(worksheet, start_pos, DATA, STYLES):
    def draw_header( cur_row, cur_col, header_data):
        for item in header_data:
            coord, text, style = item if type(item)!= list else item[0] 
            sx,sy = dec_int(coord)
            draw_text_rect(worksheet, (cur_row, cur_col), (sx, sy),  text, get_style(STYLES,style))
            if type(item)== list:
                draw_header(cur_row+sy,cur_col, item[1:])
            cur_col += sx

    start_row, start_col = xl_cell_to_rowcol(start_pos)
    worksheet.set_row(start_row+3, 64)
    draw_header(start_row, start_col, DATA)
    return start_row+3

def draw_table_body(worksheet, start_pos, DATA, attrib, STYLES):
    
    table_styles, DELTAS = attrib
    start_row, start_col = start_pos
    cur_row  = start_row 
    cur_col  = start_col
    for predm in DATA:
        for n,frag in enumerate(predm):
            stl = table_styles.get(n, table_styles['def'])
            style = get_style(STYLES,stl)
            draw_text_rect(worksheet, (cur_row, cur_col),(DELTAS[n],1), frag, style)
            cur_col += DELTAS[n]
        if len(predm[2])> 56:
            worksheet.set_row(cur_row, 22)
            
        cur_row +=1
        cur_col  = start_col
    return cur_row
    
            
        
    
 
    
