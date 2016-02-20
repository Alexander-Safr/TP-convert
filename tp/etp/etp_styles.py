def get_style_(STYLES, stl = 't11:l:_'):
    ALIGN_HOR = {'l':'left', 'r': "right", 'c':'center'}
    ALIGN_VER = {'t':'top', 'v': "vcenter", 'b':'bottom'}

    if stl not in STYLES:
        new_style = {}
        font, align, attrib, *border = stl.split(':')
        if font[0] == 't':
            new_style.update({'font_name': 'Times New Roman'})
        new_style.update({'font_size': int(font[1:])})
        if len(align)==1:
            new_style.update({'align': ALIGN_HOR[align]})
            new_style.update({'valign': "vcenter"})
        elif len(align)==2:
            new_style.update({'align':  ALIGN_HOR[align[0]]})
            new_style.update({'valign': ALIGN_VER[align[1]]})
        if 'b' in attrib:
            new_style.update({'bold': True})
        if 'i' in attrib or 'k' in attrib:
            new_style.update({'italic': True})
        if 'r'  in attrib:
            new_style.update({'rotation': 90})
        if 'w'  in attrib:
            new_style.update({'text_wrap': True})
        if border != []:
            new_style.update({'border': int(border[0])})
            
        
        new_wb_style = STYLES.add_format(new_style)
        STYLES.update({stl:new_wb_style})
    return STYLES[stl]
