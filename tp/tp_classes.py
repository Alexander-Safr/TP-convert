#v 1.1
class Predmet():
    PREDM_PARTS = 'lec', 'prac', 'lab', 'srs', "cred", 'ekz','zal','kp','kr','rgr'
   

    

    def __init__(self, code, title, semdata, smap, kafedra, srs_min = None):
        self.code    = code
        self.title   = title.strip()
        self.semdata = tuple(semdata)
        self.smap    = tuple(smap)
        self.srs_min = tuple(srs_min) if srs_min else None 
        self.part    = None
        self.cycle   = None
        self.block   = None
        self.kafedra = kafedra

    def _sum_col_semdata(self, col):
        #(36, 18, 0, 54, 3.0, None, 'ekz')
        assert 0<=col<=4, "экзамен, кргр слаживать нельзя"
        return sum(list(zip(*self.semdata))[col])

    def sum_part(self, part):
        assert part in self.PREDM_PARTS, "непрвильное название элемента предмета: "+PREDM_PARTS
        if part in ('lec', 'prac', 'lab', 'srs' , "cred"): 
            return self._sum_col_semdata(self.PREDM_PARTS.index(part))
        else:
            return (sum(self.semdata, [])).count(part)
        

    @property
    def start_sem(self):
        return self.smap.index(1)+1

    @property
    def q_sem(self):
        return sum(self.smap)

    @property
    def start_course(self):
        #int((7/2 + 0.5))
        return int((self.smap.index(1)+1)/2 +0.5)

    def is_in_sem(self,semestr):
        return bool(self.smap[semestr-1])

    def is_in_course(self,course):
        return self.is_in_sem(course*2) or self.is_in_sem(course*2-1)

    def get_sem(self,semester):
        if not self.is_in_sem(semester):
            return (0,0,0,0,0.0,0,0)
        else:
            return tuple(self.semdata[semester-self.start_sem])

    def get_course(self,course):
        return (self.code, self.title, self.get_sem(course*2-1), self.get_sem(course*2) )

    def load_all(self):
        out = 0
        for s in self.semdata:
            out += sum(s[0:4])
        return out

    def load_in_course(self,course):
        data = self.get_course(course)
        sem1 = data[-2]
        sem2 = data[-1]
        return sum(sem1[0:4])+sum(sem2[0:4])

##    def load_before(self, semester):
##        if semester - 1 <= 0:
##            return 0
##        q = sum(self.smap[:semester-1])
##
##        if q == 0:
##            return 0
##        else:
##            out = 0
##            semdata = self.semdata[:q]
##            for s in semdata:
##                out += sum(s[0:4])
##            return out
            
            
        
            

        
    def __repr__(self):
        return "{}\t{}\t{}\t{}".format(self.code, self.title, self.start_sem, self.q_sem)

