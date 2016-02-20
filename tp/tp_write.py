from .tp_classes import Predmet
#import pickle
import xlsxwriter
from xlsxwriter.utility import xl_range
#from write2xlsx import *
from .etp import title_header, write_rnp 


class dstyle(dict):
    pass

def TP_Write(tp_filename, predmets, TP_INFO ,TP_ATTRIB, TP_TITLE_TABLES):
##    with open('pred_data.pickle', 'rb') as f:
##        predmets = pickle.load(f)
##    with open('title_data2.pickle', 'rb') as f:
##        title_data = pickle.load(f)
##    with open('tp_info.pickle', 'rb') as f:
##        TP_INFO = pickle.load(f)


    ##   'level': "t11:r:bi", "domain_code" : "t11:l:bi",# 'domain_name' :"U8:AL8", "qualification" : "AR8:BA8",
    ##    'trend_code': "F9:I9", 'trend_name': "K9:AL9", 'trend': "AN9:BA9",
    ##    'spec_code': "F10:J10", 'spec_name': "L10:AL10", 'term': "AS10:BA10",
    ##    
    ##    "specialization" : "F11:AL11", "base" : "AQ11:BA11",
    ##    "teach_form" : "O12:AL12",

    TP_INFO['spec_code'] = TP_ATTRIB['trend_code'][1]
    TP_INFO['spec_name'] = TP_ATTRIB['trend_name'][1]
    #YEAR = title_data[0]['year'][1]
    #TP_INFO['year'] = YEAR
    TP_INFO['kalendar'] = TP_TITLE_TABLES['kalendar']



    def get_sem(sem_data):
        #
        c2 = 'З' if sem_data[-1] == 'zal' else 0
        c1 = 'Е' if sem_data[-1] == 'ekz' else 0
        kred = sem_data[-3] or 0
        rgr = 'ргр' if sem_data[-2] == 'rgr' else 0

        if sem_data[-2] == 'kr':
            krp = 'КР'
        elif sem_data[-2] == 'kp':
            krp = 'КП'    
        else:
            krp = 0
            
        (lec,pract,lab,srs) = sem_data[0:4]
        aud_load = sum((lec,pract,lab))
        full_load = aud_load+srs
        return full_load, aud_load, lec,pract,lab,srs, krp, rgr, kred, c1, c2       
        

    def get_course_TP(c,predmets):
        def replace0_space(raw):
            for n in range(2,len(raw)-1):
                if raw[n] == 0:
                    raw[n] = ''
            return raw
            
        kaf = "УСБЖД"
        n = 1
        out = []
        for p in predmets:
            if p.is_in_course(c):
                cur_p = p.get_course(c)
                sem_1 = get_sem(cur_p[-2])
                sem_2 = get_sem(cur_p[-1])
                all_in_year = sem_1[0]+sem_2[0]
                all_kred_in_year = sem_1[-3]+sem_2[-3]
                load_all = p.load_all()
                if c > 1:
                    last_year_load = p.load_in_course(c-1)
                else:
                    last_year_load = 0
                cur_predm = [n, p.code, p.title, all_kred_in_year, load_all, load_all , last_year_load, all_in_year ]
                cur_predm.extend(sem_1)
                cur_predm.extend(sem_2)
                cur_predm.append(kaf)
                n+=1
                out.append(replace0_space(cur_predm))
        return out


    #print(get_course_TP(1,predmets))
    workbook = xlsxwriter.Workbook(tp_filename)

    STYLES = dstyle()
    STYLES.add_format = workbook.add_format

    worksheet = workbook.add_worksheet('ТитулНП')
    title_header(worksheet, TP_ATTRIB, TP_TITLE_TABLES, STYLES)

    worksheet = workbook.add_worksheet('1 курс')
    ##make_TP(workbook, worksheet, get_course_TP(1,predmets),1)
    rnp_data = (TP_INFO, get_course_TP(1,predmets),1)
    write_rnp(worksheet, rnp_data, STYLES)
    ##
    worksheet = workbook.add_worksheet('2 курс')
    rnp_data = (TP_INFO, get_course_TP(2,predmets),2)
    write_rnp(worksheet, rnp_data, STYLES)
    ##
    worksheet = workbook.add_worksheet('3 курс')
    rnp_data = (TP_INFO, get_course_TP(3,predmets),3)
    write_rnp(worksheet, rnp_data, STYLES)
    ##
    worksheet = workbook.add_worksheet('4 курс')
    rnp_data = (TP_INFO, get_course_TP(4,predmets),4)
    write_rnp(worksheet, rnp_data, STYLES)
    #print(YEAR)
    workbook.close()
