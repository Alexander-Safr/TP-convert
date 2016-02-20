import xlrd
import re
import pickle
from .xlsutils   import *
from .tp_classes import Predmet
from .tp_info_parse import  get_tp_info
from .tp_tit_parse  import  tit_parse

def TP_Parse(tp_filename):
    SHEET_TP_NAME = "РНП"
    SHEET_TITLE_NAME = "ТитулНП"
    PREDM_DATA_START = 3



    PRED_FRAGS = ('ekz','zal','kp','kr','rgr','normativ', 'full_ld','cred',
                   'aud_ld', 'lec', 'prac', 'lab', 'srs')

    SEM_POS = ('Q','W','AC', "AI", "AO", "AU", "BA", "BG" )



    def extract_predmet(raw):

        def extract_semdata(raw):
            def fix_blank(x):
                return 0 if x == '' else x
            sem_data = []
            sem_map = []
            srs_min= []
            for s in SEM_POS:
                s = col_coord(s)
                frag = raw[s:s+6]
                #del frag[-3]
                frag = list(map(fix_blank,frag))
                srs_min.append(frag.pop(-2))# srs_min remove 
                frag[:-1] = map(int,frag[:-1])
                if sum(frag):
                    sem_data.append(frag)
                    sem_map.append(1)
                else:
                    sem_map.append(0)
                    
            return sem_data, sem_map, srs_min

        def is_cw_element(test_str,cur_sem):
            test_str = str(test_str)
            if str(cur_sem) in test_str:
                return True
            if (len(test_str)>2) and ('-' in test_str):
                (x1,x2) = test_str.split('-')
                if int(x2[0])>cur_sem>int(x1[-1]):
                    return True
            return False
                
            
        def def_control(info, cur_sem):
            if is_cw_element(info['ekz'],cur_sem):
                return 'ekz'
            elif is_cw_element(info['zal'],cur_sem):
                return 'zal'
            else:
                return None
        
        def def_krgr(info, cur_sem):
            if is_cw_element(info['rgr'],cur_sem):
                return 'rgr'
            elif is_cw_element(info['kr'],cur_sem):
                return 'kr'
            elif is_cw_element(info['kp'],cur_sem):
                return 'kp'
            else:
                return None    

        code = raw[0]
        title = raw[1]
        kafedra = raw[col_coord('BM')]
        gen_info  = {f:raw[PREDM_DATA_START+PRED_FRAGS.index(f)] for f in PRED_FRAGS}
        semdata, smap, srs_min = extract_semdata(raw)
          
        q_sem = sum(smap)
        cur_sem = smap.index(1)+1
        for frag in semdata:
            control = def_control(gen_info, cur_sem)
            krgr    = def_krgr(gen_info, cur_sem)
            cur_sem += 1
            frag.append(krgr)
            frag.append(control)

            
        #print(code,title, gen_info,'\n\t',semdata,'\n\t', smap)
        return Predmet(code, title, semdata, smap, kafedra, srs_min) 
        
        
        


        
        
            
        
            
            
            
            

            


    tp = xlrd.open_workbook(tp_filename,formatting_info=False)







      

    def test_row_content(r):
        if isinstance(r,(int,float)) or r =="":
            return False
        ptrn_pred_code = r"(?:ГСЕ|МПН|ПП)(?:\.|\s+)\d\.\d+" # шаблон для кода предмета
        if type(r) == str and re.search(ptrn_pred_code,r.strip()):
            return "predmet"
            
        elif "НОРМАТИВНІ НАВЧАЛЬНІ ДИСЦИПЛІНИ" in r:
            status['part'] = "НОРМАТИВ"
            status['cycle']= None
            status['block']= None
            
        elif "ВИБІРКОВІ НАВЧАЛЬНІ ДИСЦИПЛІНИ" in r:
            status['part'] = "ВАРІАТИВ"
            status['cycle']= None
            status['block']= None
        elif "Гуманітарні та соціально-економічні дисципліни" in  r:
            status['cycle']= "ГСЕ"
            status['block']= None

        elif "Дисципліни математичної, природничо-наукової (фундаментальної) підготовки" in  r:
            status['cycle']= "МПН"
            status['block']= None

        elif "Дисципліни професійної і практичної підготовки" in  r:
            status['cycle']= "ПП"
            status['block']= None

        elif ("Професійна підготовка" in r) and status['cycle']== "ПП" :
            status['block']= "ПрофП"
        elif ("Практична підготовка" in r) and status['cycle']== "ПП" and status['block']== "ПрофП":
            status['block']= "ПрактП"
        
        elif "Дисципліни самостійного вибору навчального закладу" in  r:
            status['cycle']= "СВНЗ"
            status['block']= None

        elif "Дисципліни вільного вибору студента" in  r:
            status['cycle']= "ВВС"
            status['block']= None
    ##    else:
    ##        print("!!!!", r, r2)

        elif ("Блок А" in r) and status['cycle']== "ВВС" :
            status['block']= "Блок А"
        elif ("Блок Б" in r) and status['cycle']== "ВВС" :
            status['block']= "Блок Б"
        elif ("Блок В" in r) and status['cycle']== "ВВС" :
            status['block']= "Блок В"

        
        
    #----MAIN------
    predm_map = []
    stp = tp.sheet_by_name(SHEET_TP_NAME)
    #nrows = stp.nrows
    raw_col0,raw_col1 = stp.col_values(0),stp.col_values(1)
    status = {'part':None, 'cycle':None, 'block':None }


    for i,d in enumerate(raw_col0):
        s = test_row_content(d) # проверка, что в строке     
        if s == "predmet":
            #print(i, d, raw_col1[i],  status['part'], status['cycle'],  status['block'])
            predm_map.append((i, d, raw_col1[i],  status['part'], status['cycle'],  status['block']))
            #print(predm_map[-1])

    ##
    ##for p in predm_map:
    ##    print(p)
    ####    
    ##0/0        
    predmets = []
    for pred in predm_map:
        raw = stp.row_values(pred[0])
        
        predmets.append(extract_predmet(raw))
       
        predmets[-1].part  = pred[-3]
        predmets[-1].cycle = pred[-2]
        predmets[-1].block = pred[-1]

##    n = 1
##
##    def add_predm(predm):
##        for i in range(len(psum)):
##            psum[i]+=predm[i]
##        
##
##    sem = 1
##    psum = [0,0,0,0,0.0]
##    for p in predmets:
##        if p.is_in_sem(sem):
##            print(n, p.get_sem(sem))
##            add_predm(p.get_sem(sem))
##            n+=1
##    print(psum)
    stit = tp.sheet_by_name(SHEET_TITLE_NAME)        
    TP_ATTRIB, TP_TITLE_TABLES  = tit_parse(stit)
    TP_INFO = get_tp_info(stp)
    TP_INFO['year'] = TP_ATTRIB['year'][1]
    return predmets, TP_INFO, TP_ATTRIB, TP_TITLE_TABLES 

##    #print(get_tp_info(stp))
##    with open('tp_info.pickle', 'wb') as f:
##        pickle.dump(get_tp_info(stp), f)
##    with open('pred_data.pickle', 'wb') as f:
##        pickle.dump(predmets, f)


