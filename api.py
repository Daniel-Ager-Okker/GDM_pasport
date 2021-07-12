import os, ctypes, json
class API:
    def __init__(self):
        with os.add_dll_directory(os.getcwd()+"\\api_pebi"):
            self.pebidll = ctypes.CDLL('api_pebi_dll-3258-c10b7608.dll')
        self.pebidll.set_json_data_pod.argtypes = [ctypes.c_char_p]
        class pebi_string_pod(ctypes.Structure):
            #указатель на zero-terminated строку
            _fields_=[("s",ctypes.c_char_p), ("size",ctypes.c_int), ("capacity",ctypes.c_int)]
        self.pebidll.set_json_data_pod.restype = pebi_string_pod
        self.pebidll.openlog.argtypes = [ctypes.c_char_p]
        self.pebidll.openlog.restype = None
        
    def sjd(self,v):
        s = json.dumps(v)
        b_s = s.encode('utf-8')
        b_s_pod = self.pebidll.set_json_data_pod(b_s)
        s_pod = b_s_pod.s.decode("utf-8")
        data = json.loads(s_pod)
        return data

    #загрузить модель
    def load_model(self,path): #загрузка гд модели
        model_id=self.sjd({'action':'new_model'})['model_id']
        self.sjd({'model_id':model_id, 'action':'new_gdm_model'})
        self.sjd({'model_id':model_id, 'action':'load_gdm_model','name':path, 'request_output':True})
        return model_id #возвращает id модели
    
    #узнать размеры сетки
    def get_dimens(self,model_id):
        dims=self.sjd({'model_id':model_id,'action':'get_dimens'})
        return (dims) #возвращает размеры по осям Х,Y,Z
    
    #инициализировать модель
    def init(self,model_id): #на вход принимает id модели
        self.sjd({'model_id':model_id, 'action':'rewind'})
        calcmodel = self.sjd({'model_id':model_id, 'action':'create_calcmodel'})
        initgdm = self.sjd({'model_id':model_id, 'action':'init_gdm_model'})
        return initgdm
    
    #данные с слова EQUIL    
    def get_equil(self,model_id, region_number):
        data_eq=self.sjd({'model_id':model_id, 'action':'get_equil','region':int(region_number)})
        return data_eq
    
    #число регионов PVT
    def get_number_of_pvt_regions(self,model_id):
        number=self.sjd({'model_id':model_id, 'action':'get_number_of_pvt_tables'})
        return number
    
    #число регионов инициализации
    def eq_number(self,model_id):
        massiv=self.sjd({'model_id':model_id,'action':'get_number_of_equil_tables'})
        return massiv
    
    #получить табличку PBVD
    def get_pbvd(self,model_id,region_number):
        value=self.sjd({'model_id':model_id,'action':'get_pbvd','region':int(region_number)})
        return value
    
    #получить табличку PVTO
    def pvto(self,model_id,region_number):
        table=self.sjd({'model_id':model_id,'action':'get_pvto','region':int(region_number)})
        return table
    
    #получить табличку PVDG
    def get_pvdg(self,model_id,region_number):
        value=self.sjd({'model_id':model_id,'action':'get_pvdg','region':int(region_number)})
        return value
    
    #получить табличку PVCDO
    def get_pvcdo(self,model_id,region_number):
        value=self.sjd({'model_id':model_id,'action':'get_pvcdo','region':int(region_number)})
        return value
    
    #число табличек ОФП
    def get_number_of_relperm_tables(self,model_id,region_number):
        value=self.sjd({'model_id':model_id,'action':'get_number_of_relperm_tables','region':int(region_number)})
        return value
    
    #получить сжимаемость породы
    def get_rock(self,model_id,region_number):
        rock=self.sjd({'model_id':model_id,'action':'get_rock','region':int(region_number)})
        return rock
    
    #получить значение газосодержания
    def get_rs(self,model_id):
        value=self.sjd({'model_id':model_id, 'action':'get_rs'})
        return value
    #получить табличку RSVD
    def get_rsvd(self,model_id,region_number):
        value=self.sjd({'model_id':model_id,'action':'get_rsvd','region':int(region_number)})
        return value
    
    #число скважин в модели
    def get_well_list(self,model_id):
        massiv=self.sjd({'model_id':model_id,'action':'get_well_list'})
        return massiv
    
    #функция поиска Rs по табличке PVT в случае отутствия слова RS/RSVD
    def get_Rs_due_pbub(self,table,pbvd_pressure): #table - это табличка PVTO
        k=[] #массив позиций со сменами
        Rs=table[0][0]
        for i in range(1,len(table)):
            if table[i][0]!=table[i-1][0]:
                k.append(i)
        for i in range(0,len(table)):
            if (i in k) and (table[i][1]==pbvd_pressure):
                Rs=table[i][0]
        return Rs
    
    #информация по воде
    def get_pvtw(self,model_id,region_number):
        value=self.sjd({'model_id':model_id,'action':'get_pvtw','region':int(region_number)})
        return value
    
    def bub_pres_find(self,table,Rs): #функция поиска давления насыщения по табличке PVTO
        for el in table:
            if el[0]==Rs:
                bub_p=el[1]
                break #нашли первую такую строчку и вышли из цикла
        return bub_p
    
    #функция поиска Bo по табличке PVT в случае отсутствия лова PVDO
    def boil_find(self,table,Rs,pbvd_pressure): #table - табличка PVTO; Rs - газосодержание
        Bo=1.0
        for i in range(0,len(table)):
            if (table[i][1]==pbvd_pressure) and (table[i][0]==Rs): #отыскали нужную строчку
                Bo=table[i][2]
        return Bo
    
    #функция поиска вязкости по табличке PVT
    def visc_find(self,table,Rs,pbvd_pressure): #table - табличка PVTO
        visc=1
        for i in range(0,len(table)):
            if (table[i][1]==pbvd_pressure) and (table[i][0]==Rs): #отыскали нужную строку
                visc=table[i][3]
        return visc
    
    def get_ntg(self,model_id):
        value=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'NTG'})
        return value
    
    def get_permx(self,model_id):
        value=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'PERMX'})
        return value
    
    def get_permy(self,model_id):
        value=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'PERMY'})
        return value
    
    def get_permz(self,model_id):
        value=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'PERMZ'})
        return value
    
    def get_poro(self,model_id):
        value=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'PORO'})
        return value
    
    def get_actnum(self,model_id):
        value=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'ACTNUM'})
        return value
    
    def get_fipnum(self,model_id):
        value=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'FIPNUM'})
        return value
    
    def get_multpv(self,model_id):
        value=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'MULTPV'})
        return value
    
    def get_soil(self,model_id):
        value=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'SOIL'})
        return value
    
    def get_swat(self,model_id):
        value=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'SWAT'})
        return value
    
    def get_sgas(self,model_id):
        value=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'SGAS'})
        return value
        
    def get_densities(self,model_id,region):
        value=self.sjd({'model_id':model_id,'action':'get_densities','region':int(region)})
        return value
    
    def get_events(self,model_id):
        return self.sjd({'model_id':model_id, 'action':'get_event'})
    
    def get_depth(self,model_id):
        ae=self.sjd({'model_id':model_id,'action':'array_editor','formula':'TMP1=DEPTH'})
        values=self.sjd({'model_id':model_id,'action':'get_arr_values','name':'TMP1'})
        return values