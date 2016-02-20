from tp import TP_Parse, TP_Write


#predmets, TP_INFO ,TP_ATTRIB, TP_TITLE_TABLES = TP_Parse('TP2.xls')
#TP_Write('res.xlsx', predmets, TP_INFO ,TP_ATTRIB, TP_TITLE_TABLES)

TP_Write('result.xlsx', *TP_Parse('in_TP.xls'))


