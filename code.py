'''print('Введите количество объектов (пластов/пропластков)')
n_obj=int(input())
print('Для каждого объекта введите номер начального слоя в модели и конечного')
z_begins=[] #массив стартовых слоёв
z_ends=[] #массив конечных слоёв
for i in range(0,n_obj):
    print(f"Номер начального слоя для {i+1}-того объекта")
    strt=int(input()) #зачитали
    z_begins.append(strt) #сохранили
    print(f"Номер конечного слоя для {i+1}-того объекта")
    fnsh=int(input()) #зачитали
    z_ends.append(fnsh) #сохранили'''
import os
print(os.getcwd())

#работа с API
import ctypes
import os
with os.add_dll_directory(os.getcwd()):
    pebidll = ctypes.CDLL('api_pebi_dll-3241-b330d55e.dll')
pebidll.set_json_data_pod.argtypes = [ctypes.c_char_p]
class pebi_string_pod(ctypes.Structure):
    #указатель на zero-terminated строку
    _fields_ = [("s", ctypes.c_char_p),
    #размер строки без конечного нуля
    ("size", ctypes.c_int),
    #выделенная память
    ("capacity", ctypes.c_int)]

pebidll.set_json_data_pod.restype = pebi_string_pod
pebidll.openlog.argtypes = [ctypes.c_char_p]
pebidll.openlog.restype = None
    
import json
def sjd(v):
    s = json.dumps(v)
    b_s = s.encode('utf-8')
    b_s_pod = pebidll.set_json_data_pod(b_s)
    s_pod = b_s_pod.s.decode("utf-8")
    data = json.loads(s_pod)
    return data

#функция загрузки гд модели
def load_model(path): #на вход принимает имя модели
    model_id=sjd({'action':'new_model'})["model_id"]
    sjd({'model_id':model_id, 'action':'new_gdm_model'})
    sjd({'model_id':model_id, 'action':'load_gdm_model','name':path, 'request_output':True})
    return model_id #возвращает id модели

#функция инициализации гд модели
def init(model_id): #на вход принимает id модели
    sjd({'model_id':model_id, 'action':'rewind'})
    calcmodel = sjd({'model_id':model_id, 'action':'create_calcmodel'})
    initgdm = sjd({'model_id':model_id, 'action':'init_gdm_model'})
    return initgdm

#функция получения куба песчанистости
def ntg(model_id):
    d=sjd({'model_id':model_id,'action':'get_arr_values','name':'NTG'})
    return d

#функция получения размеров модели
def get_dimens(model_id):
    dims=sjd({'model_id':model_id,'action':'get_dimens'})
    return (dims) #возвращает размеры по осям Х,Y,Z

#функция получения куба проницаемости по Х
def get_ntg(model_id):
    values=sjd({'model_id':model_id, 'action':'get_arr_values','name':'NTG'})
    return values

#функция получения данных по слову equil
def get_equil(model_id,region_number):
    data_eq=sjd({'model_id':model_id, 'action':'get_equil','region':int(region_number)})
    return data_eq

#функция получения данных по сжимаемости
def get_rock(model_id):
    rock=sjd({'model_id':model_id,'action':'get_rock'})
    return rock

#функция получения данных по количеству регионов PVT
def get_number_of_pvt_regions(model_id):
    number=sjd({'model_id':model_id, 'action':'get_number_of_pvt_tables'})
    return number

#функция получения данных по плотности
def get_pvt(model_id):
    pvt=sjd({'model_id':model_id,'action':'get_pvt','region':1})
    return pvt

#функции получения данных по газосодержанию
import math #нужна библиотека для округления
def get_rs(model_id):
    value=sjd({'model_id':model_id, 'action':'get_rs'})
    return value
def get_rsvd(model_id,region_number):
    value=sjd({'model_id':model_id,'action':'get_rsvd','region':int(region_number)})
    return value

#функция получения инф-ции по числу слоев
def layers(model_id):
    number=sjd({'model_id':model_id,'action':'get_number_of_layers'})
    return number

#функция получения таблички PVTO
def pvto(model_id,region_number):
    table=sjd({'model_id':model_id,'action':'get_pvto','region':int(region_number)})
    return table

#функция получения значения давления насыщения
def get_pbvd(model_id,region_number):
    value=sjd({'model_id':model_id,'action':'get_pbvd','region':int(region_number)})
    return value

#функция линейной интерполяции значения по таблице формата PBVD/RSVD
def interpol(table,h_ref): #прнимает на вход СПИСОК и опорную глубину
    value=0
    if table==None:
        return value
    else:
        depths=[] #создали массив с глубинами
        values=[] #массив с значенями искомой величины
        for i in range(0,len(table)):
            if (i % 2 == 0):
                depths.append(table[i]) #заполянем глубины
            else:
                values.append(table[i]) #величины
        #находим диапазон
        H_bottom=0
        t,b=0,0
        for d in range(1,len(depths)):
            if (depths[d-1] < h_ref) and (depths[d] > h_ref):
                H_bottom=depths[d-1]
                b=d-1
                t=d
        value=values[b]+float((h_ref-H_bottom)*(values[t]-values[b]))/(depths[t]-depths[b])
        return value

#функция поиска Rs по табличке PVT в случае отутствия слова RS/RSVD
def rs_find(table,pbvd_pressure): #table - это табличка PVTO
    k=[] #массив позиций со сменами
    Rs=table[0][0]
    for i in range(1,len(table)):
        if table[i][0]!=table[i-1][0]:
            k.append(i)
    for i in range(0,len(table)):
        if (i in k) and (table[i][1]==pbvd_pressure):
            Rs=table[i][0]
    return Rs
 
def bub_pres_find(table,Rs): #функция поиска давления насыщения по табличке PVTO
    for el in table:
        if el[0]==Rs:
            bub_p=el[1]
            break #нашли первую такую строчку и вышли из цикла
    return bub_p
            
#функция поиска Bo по табличке PVT в случае отсутствия лова PVDO
def boil_find(table,Rs,pbvd_pressure): #table - табличка PVTO; Rs - газосодержание
    Bo=1.0
    for i in range(0,len(table)):
        if (table[i][1]==pbvd_pressure) and (table[i][0]==Rs): #отыскали нужную строчку
            Bo=table[i][2]
    return Bo

#функция поиска вязкости по табличке PVT
def visc_find(table,Rs,pbvd_pressure): #table - табличка PVTO
    visc=1
    for i in range(0,len(table)):
        if (table[i][1]==pbvd_pressure) and (table[i][0]==Rs): #отыскали нужную строку
            visc=table[i][3]
    return visc


#функция получения списка скважин
def get_well_list(model_id):
    massiv=sjd({'model_id':model_id,'action':'get_well_list'})
    return massiv

#ф-ция получения кол-ва регионов ОФП
def get_relperm_number(model_id):
    massiv=sjd({'model_id':model_id,'action':'get_number_of_relperm_tables'})
    return massiv

#ф-ция получения кол-ва регионов PVT
def pvt_number(model_id):
    massiv=sjd({'model_id':model_id,'action':'get_number_of_pvt_tables'})
    return massiv

#ф-ция получения кол-ва регионов инициализации
def eq_number(model_id):
    massiv=sjd({'model_id':model_id,'action':'get_number_of_equil_tables'})
    return massiv

#хз
def get_interp_data(model_id):
    value=sjd({'model_id':model_id,'action':'get_interp_data'})
    return value

#переходим в ту папку, где лежит модель
#print('Введите имя папки, где находится модель')
model_folder='мнимый пилот' #ЗДЕСЬ ДОЛЖЕН БЫТЬ INPUT()
current_folder=os.getcwd() #определили путь той папки, где находистя исполняющий скрипт
os.chdir(os.path.join(current_folder,model_folder)) #перешли в папку с моделью
model_folder_files=os.listdir() #составили список файлов в папке с моделью

for file in model_folder_files:
    if ('.DATA' in file) or ('.data' in file):
        model_name=file #запомнили имя модели

model_id=load_model(model_name) #загрузили модель и запомнили её id

print('Каким количеством пропластков представлена модель?')
n=input()
eq_regions=eq_number(model_id)['EQUIL'] #кол-во регионов инициализации 

reference_depth=[] #значения опорных глубин для каждого региона инициализации
reference_pressure=[] #значения пластовых давлений (на опорной глубине)
VNK=[] #глубины ВНК
GNK=[] #глубины ГНК
for i in range(0,eq_number(model_id)['EQUIL']):
    reference_depth.append(get_equil(model_id, i)['href']) #заполняем для соответствуюещго региона инициализации
    reference_pressure.append(get_equil(model_id,i)['ref_pres']) #давления для каждого региона
    VNK.append(get_equil(model_id,i)['hwoc']) #глубины ВНК
    GNK.append(get_equil(model_id,i)['hgoc']) #глубины ГНК

rock=get_rock(model_id)['compr'] #значение сжимаемости породы

pvt_regions=pvt_number(model_id)['PVTO'] #кол-во регионов PVT

#в модели может быть либо RSVD, либо PBVD
PBVD_tables_number=eq_number(model_id)['PBVD'] #число табличек PBVD
RSVD_tables_number=eq_number(model_id)['RSVD'] #число табличек RSVD
bub_pressure=[] #значения давления насыщения для каждого региона
Rs=[] #значения газосодержания
if (PBVD_tables_number != 0): #если у нас есть ключевое слово PBVD
    for i in range(0,eq_number(model_id)['PBVD']):
        table_pbvd=get_pbvd(model_id,i)['table'] #взяли табличку для i-того региона
        bub_pressure.append(interpol(table_pbvd,reference_depth[i])) #взяли значения давлений насыщения
        for j in range(0,pvt_regions): #СОВПАДАЮТ НУМЕРАЦИИ РЕГИОНОВ PVT И ИНИЦИАЛИЗАЦИИ
            pvto_table=pvto(model_id,j)['table']
            Rs.append(rs_find(pvto_table,bub_pressure[i])) #и ищем Rs по табличке PVT
else: #если у нас его нет
    for i in range(0,eq_number(model_id)['RSVD']): #значит есть слово RSVD
        table_rsvd=get_rsvd(model_id,i)['table'] #забираем Rs с слова
        Rs.append(interpol(table_rsvd, reference_depth[i])) #интерполируем
        for j in range(0,pvt_regions): 
            pvto_table=pvto(model_id,j)['table']
            bub_pressure.append(bub_pres_find(pvto_table,Rs[i])) #по PVTO ищем давление насыщения
            

'''for i in range(0,eq_regions):
    reference_depth.append(get)


Bo=boil_find(pvto(model_id)['table'],Rs,bub_pressure) #объемный коэффициент 
viscosity=visc_find(pvto(model_id)['table'],Rs,bub_pressure)

dimens=get_dimens(model_id)

wells=len(get_well_list(model_id)) #число скважин
ofp_regions=get_relperm_number(model_id)['SWOF'] #кол-во регионов ОФП



#print(get_interp_data(model_id))

#init(model_id) #инициализировали модель, чтобы зачитать кубы



#выгрузка в excel
import xlsxwriter
f_out=xlsxwriter.Workbook('report.xlsx') #создали результирующйи сводный файл
sheet_out=f_out.add_worksheet('ГФХ') #создали лист ГФХ
#формат ячеек-заголовков
format1=f_out.add_format()
format1.set_bold() #жирный шрифт
format1.set_align('center') #положение по центру
#формат ячеек-незаголовков
format2=f_out.add_format()
format2.set_align('center') #положение по центру
#делаем шапку
sheet_out.write(0,0,'Параметр',format1)
sheet_out.write(0,1,'Объект 1',format1)
sheet_out.write(1,0,'Средняя глубина залегания кровли, м',format2)
sheet_out.write(2,0,'Абсолютная отметка ВНК, м',format2)
sheet_out.write(3,0,'Абсолютная отметка ГНК, м',format2)
sheet_out.write(4,0,'Опорная глубина, м',format2)
sheet_out.write(5,0,'Пластовое давление, бар',format2)
sheet_out.write(6,0,'Объемный коэффициент нефти, д. ед.',format2)
sheet_out.write(7,0,'Давление насыщения, бар',format2)
sheet_out.write(8,0,'Газосодержание, м3/т',format2)
sheet_out.write(9,0,'Вязкость нефти в пл. условиях, сПз',format2)
sheet_out.write(10,0,'Сжимаемость породы, 1/бар',format2)
#форматирование шапки
sheet_out.set_column(0,0,35) #длина нулевого (первого) столбца по макс. значению
#заполняем значениями
sheet_out.write(2,1,VNK,format2)
sheet_out.write(3,1,GNK,format2)
sheet_out.write(4,1,reference_depth,format2)
sheet_out.write(5,1,reference_pressure,format2)
sheet_out.write(6,1,Bo,format2)
sheet_out.write(7,1,bub_pressure,format2)
sheet_out.write(8,1,Rs,format2)
sheet_out.write(9,1,viscosity,format2)
sheet_out.write(10,1,rock,format2)
#лист "описание гд модели"
sheet_out_2=f_out.add_worksheet('гд модель')
#шапка
sheet_out_2.write(0,0,'Параметр',format1)
sheet_out_2.write(0,1,'Значение',format1)
sheet_out_2.write(1,0,'Размерность сетки',format2)
sheet_out_2.write(2,0,'Общее количество ячеек',format2)
sheet_out_2.write(3,0,'Количество активных ячеек',format2)
sheet_out_2.write(4,0,'Количество скважин',format2)
sheet_out_2.write(5,0,'Количество регионов инициализации',format2) #колв-о регионов equil'ей
sheet_out_2.write(6,0,'Количество регионов ОФП',format2)
sheet_out_2.write(7,0,'Количество регионов PVT',format2)
sheet_out_2.write(8,0,'Модель PVT свойств',format2)
sheet_out_2.write(9,0,'Термическая опция',format2)
#форматирование шапки
sheet_out_2.set_column(0,0,55) #длина нулевого (первого) столбца по макс. значению
sheet_out_2.set_column(1,0,25) #длина второго столбца по макс. значению
#значения
sheet_out_2.write(1,1,f"{dimens['nx']}x{dimens['ny']}x{dimens['nz']}",format2)
sheet_out_2.write(4,1,wells,format2)
sheet_out_2.write(5,1,eq_regions,format2)
sheet_out_2.write(6,1,ofp_regions,format2)
sheet_out_2.write(7,1,pvt_regions,format2)
f_out.close()'''