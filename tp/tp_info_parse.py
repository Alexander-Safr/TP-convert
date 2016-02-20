from xlsxwriter.utility import xl_range, xl_cell_to_rowcol as coord

def get_tp_info(worksheet):
    TP_MAP = {
        'decan_pos'  :"B166" , 
        'zavkav_pos' : "B169",
        'decan_name'  :"C166" , 
        'zavkav_name'   : "C169",
        'prorector_pos' :"N166" , 
        'HMB_pos'       : "N169",
        'prorector_name' :"Y166" , 
        'HMB_name'       : "Y169",
    }
    TP_INFO = {}
    for key in TP_MAP:
        TP_INFO[key] = worksheet.cell(*coord(TP_MAP[key])).value

    return TP_INFO


    
