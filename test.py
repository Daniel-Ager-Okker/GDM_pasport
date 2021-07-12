from PyQt5 import QtCore, QtGui, QtWidgets
from form import Ui_MainWindow
import os, api, xlsxwriter, math
#import matplotlib.pyplot as plt

def unique_num(array): #фнукция поиска числа уникальных значений
    unique=[]
    for e in range(1,len(array)):
        if array[e] in unique:
            pass
        else:
            unique.append(array[e])
    return len(unique)

def model_type(number_of_PVTO_tables):
    if number_of_PVTO_tables==0:
        PVT_type='Dead oil'
    else:
        PVT_type='Live oil'
    return PVT_type

def average(list_data): #функция нахождения среднего арифметического; на вход принимает список
    summa=0
    for i in range(0,len(list_data)):
        summa=summa+list_data[i]
    avg_value=summa/len(list_data)
    return avg_value

#функция поиска среднего значения параметра данного слоя (нужно для ср. h залегания кровли)
def average_layer_param(number_of_object,object_data,layer_number,dimension,PARAM):
    depth_data={}
    depth_data['obj_name']=object_data[number_of_object]['obj_name'] #имя объекта
    depth_data['layer_number']=layer_number #номер слоя
    #пояснение
    #первая ячейка в слое - это (n-1)*nx*ny, где n - номер слоя
    start=dimension['nx']*dimension['ny']*(int(layer_number)-1)
    #последняя ячейка в слое - это n*nx*ny-1 
    finish=int(layer_number)*dimension['nx']*dimension['ny'] #но тут мы не будем вычитать единичку, так как в срезе будем бежать ДО finish
    values=PARAM[start:finish] #делаем нужный срез из общей информации по значениями
    depth_data['average']=average(values) #нашли 
    return depth_data

#функция определения средней общей и эффективной толщин объекта
def avg_thick(number_of_objects,objects_data,dimension,PARAM): #в роли PARAM выступает толщина
    thick=[] #список словарей по каждому объекту
    for i in range(0,number_of_objects):
        thick_data={} #словарь с значениями
        thick_data['obj_name']=objects_data[i]['obj_name'] #имя
        start_top=(int(objects_data[i]['top'])-1)*dimension['nx']*dimension['ny']
        end_top=int(objects_data[i]['top'])*dimension['nx']*dimension['ny']
        thick_data['values_top']=PARAM[start_top:end_top] #срезали значения
        start_down=(int(objects_data[i]['down'])-1)*dimension['nx']*dimension['ny']
        end_down=int(objects_data[i]['down'])*dimension['nx']*dimension['ny']
        thick_data['values_down']=PARAM[start_down:end_down] #срез по значениям
        thickness=[] #массив с значениями толщин
        for l in range(0,len(thick_data['values_down'])):
            thickness.append(thick_data['values_down'][l]-thick_data['values_top'][l])
        thick_data['avg_thickness']=average(thickness)
        thick.append(thick_data)
    return thick



def get_arr_layer(number_of_objects,objects_data,dimension,PARAM): #на вход: число объектов разработки; инф-ция оп ним; размерны сетки; нужный параметр
    obj_param=[] #список значений нужного параметра для каждого из объектов
    for i in range(0,number_of_objects):
        param_obj_data={} #значения параметров самих
        param_obj_data['obj_name']=objects_data[i]['obj_name'] #имя объекта
        start=dimension['nx']*dimension['ny']*(int(objects_data[i]['top'])-1) #1-АЯ (НУЛЕВАЯ) ячейка слоя
        finish=dimension['nx']*dimension['ny']*(int(objects_data[i]['down'])) #последняя ячейка слоя
        param_obj_data['values']=PARAM[start:finish] #срезами разделяем ячейки по слоям соответствующих объектов
        param_obj_data['average']=average(param_obj_data['values']) #вызвали выше описанную функцию для нахождения среднего
        obj_param.append(param_obj_data)
    return obj_param

def bar_graph_data(list_data): #принимает на вход список значений
    X=[] #в роли икс у нас будет выступать частота встречания
    Y=[] #в роли игрек - само значение
    #сначала ищем только уникальные значения
    for el in list_data:
        if el in Y:
            pass
        else:
            Y.append(el)
    for i in range(0,len(Y)):
        k=0 #число встречаний
        for j in range(0,len(list_data)):
            if Y[i]==list_data[j]:
                k+=1
        X.append(k) #добавили число встречаний
    dict_out={}
    dict_out['X']=X
    dict_out['Y']=Y
    return dict_out #возращает словарь из двух массивов


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
        return value #возвращает проинтерполированное значение
    
        

class mywindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.btnClicked)
        self.ui.tableWidget
        self.ui.progressBar
        self.ui.pushButton_2.clicked.connect(self.btnClicked_2)
        
    def btnClicked(self):
        obj_number=int(self.ui.lineEdit.text())
        QtWidgets.QMessageBox.information(self, 'Важно',
            "Заполните, пожалуйста, таблицу по объектам",
            QtWidgets.QMessageBox.Ok)
        self.ui.tableWidget.setRowCount(obj_number)
        
    def btnClicked_2(self):
        pebi=api.API() #обозначили переменную класса API
        obj_number=int(self.ui.lineEdit.text())
        #проверяем, заполнена ли таблица полностью
        exit=False
        for i in range(0,obj_number):
            for j in range(0,5):
                if self.ui.tableWidget.item(i,j)==None:
                    exit=True
                    QtWidgets.QMessageBox.information(self,'Важно',
                    "Заполните, пожалуйста, таблицу полностью",
                    QtWidgets.QMessageBox.Ok)
                    break
            if exit:
                break
            else:
                pass
        names=['obj_name','top','down','PVT_region','eq_region','fipnum_region']
        data=[] #список словарей для каждого объекта
        for i in range(0,obj_number):
            dict_info={}
            j=0
            for name in names:
                dict_info[name]=self.ui.tableWidget.item(i,j).text()
                j+=1
            data.append(dict_info)
        
        #далее нужно учесть, что нумерация по регионам идёт с нуля в программе
        #а у обычных пользователей - с единицы
        for dict_ in data:
            for key in dict_:
                if key=='PVT_region':
                    dict_[key]=str(int(dict_[key])-1)
                if key=='eq_region':
                    dict_[key]=str(int(dict_[key])-1)
                if key=='fipnum_region':
                    dict_[key]=str(int(dict_[key])-1)
        
        #print(data) #дебажный принт
        gdm_folder=self.ui.lineEdit_2.text()
        current_folder=os.getcwd() #определили путь той папки, где находистя исполняющий скрипт
        os.chdir(os.path.join(current_folder,gdm_folder)) #перешли в папку с моделью
        model_folder_files=os.listdir() #составили список файлов в папке с моделью

        for file in model_folder_files:
            if ('.DATA' in file) or ('.data' in file):
                model_name=file #запомнили имя модели

        model_id=pebi.load_model(model_name) #загрузили модель и запомнили её id
        
        
        params_from_equil=['href','ref_pres','hwoc','hgoc']
        data_from_equil=[] #инофрмация, получаемая со слова EQUIL
        for i in range(0,obj_number):
            info_from_equil={}
            for element in params_from_equil:
                info_from_equil['obj_name']=data[i]['obj_name'] #имя объекта разработки
                info_from_equil[element]=pebi.get_equil(model_id, data[i]['eq_region'])[element]
            data_from_equil.append(info_from_equil)
        #print(data_from_equil) #дебажынй принт
        EQ_number=pebi.eq_number(model_id)['EQUIL'] #число регионов инициализации/равновесия
        
        
        PBVD_tables_number=pebi.eq_number(model_id)['PBVD'] #число таблиц PBVD
        RSVD_tables_number=pebi.eq_number(model_id)['RSVD'] #число таблиц RSVD
        PVCDO_tables=pebi.get_number_of_pvt_regions(model_id)['PVCDO'] #число таблиц PVCDO
        PVDO_tables=pebi.get_number_of_pvt_regions(model_id)['PVDO'] #число таблиц PVDO
        
        
        PVT_data=[] #информация по PVTям
        if PBVD_tables_number==0 and RSVD_tables_number==0: #если нет слов PBVD и RSVD
            P_bub=None #давление насыщения
            if PVCDO_tables!=0: #если свойства PVT заданы через PVCDO
                data_from_PVCDO=[] #информация с слова PVCDO
                #ВСЁ при пластовом давлении!!!
                for i in range(0,obj_number):
                    info_from_PVCDO={}
                    info_from_PVCDO['obj_name']=data[i]['obj_name'] #имя объекта
                    info_from_PVCDO['B_oil']=pebi.get_pvcdo(model_id,data[i]['PVT_region'])['table'][1] #объемник
                    info_from_PVCDO['oil_compr']=pebi.get_pvcdo(model_id,data[i]['PVT_region'])['table'][2] #сжимаемость нефти
                    info_from_PVCDO['visc']=pebi.get_pvcdo(model_id,data[i]['PVT_region'])['table'][3] #вязкость
                    info_from_PVCDO['rock']=pebi.get_rock(model_id,data[i]['PVT_region'])['compr'] #сжимаемость породы
                    data_from_PVCDO.append(info_from_PVCDO)
                PVT_data=data_from_PVCDO #тогда у нас PVT берутся с PVCDO
        else:
            if PBVD_tables_number!=0: #если у нас есть ключевое слово PBVD
                data_from_PVTO=[] #информация по PVTO
                for i in range(0,obj_number):
                    info_from_PVTO={}
                    info_from_PVTO['obj_name']=data[i]['obj_name'] #имя объекта
                    info_from_PVTO['bub_pres']=interpol(pebi.get_pbvd(model_id,data[i]['PVT_region'])['table'],data_from_equil[i]['href'])
                    info_from_PVTO['Rs']=pebi.get_Rs_due_pbub(pebi.pvto(model_id,data[i]['PVT_region'])['table'],info_from_PVTO['bub_pres'])
                    info_from_PVTO['B_oil']=pebi.boil_find(pebi.pvto(model_id,data[i]['PVT_region'])['table'], info_from_PVTO['Rs'], info_from_PVTO['bub_pres'])
                    info_from_PVTO['visc']=pebi.visc_find(pebi.pvto(model_id,data[i]['PVT_region'])['table'], info_from_PVTO['Rs'], info_from_PVTO['bub_pres'])
                    info_from_PVTO['rock']=pebi.get_rock(model_id,data[i]['PVT_region'])['compr'] #сжимаемость породы
                    data_from_PVTO.append(info_from_PVTO)
                PVT_data=data_from_PVTO #тогда у нас PVT с таблицы PVTO
            elif RSVD_tables_number!=0: #если у нас есть ключевое слово RSVD
                data_from_PVTO=[] #информация по PVTO
                for i in range(0,obj_number):
                    info_from_PVTO={}
                    info_from_PVTO['obj_name']=data[i]['obj_name'] #имя объекта
                    info_from_PVTO['Rs']=interpol(pebi.get_rsvd(model_id,data[i]['PVT_region'])['table'],data_from_equil[i]['href'])
                    info_from_PVTO['bub_pres']=pebi.bub_pres_find(pebi.pvto(model_id,data[i]['PVT_region'])['table'], info_from_PVTO['Rs'])
                    info_from_PVTO['B_oil']=pebi.boil_find(pebi.pvto(model_id,data[i]['PVT_region'])['table'], info_from_PVTO['Rs'], info_from_PVTO['bub_pres'])
                    info_from_PVTO['visc']=pebi.visc_find(pebi.pvto(model_id,data[i]['PVT_region'])['table'], info_from_PVTO['Rs'], info_from_PVTO['bub_pres'])
                    info_from_PVTO['rock']=pebi.get_rock(model_id,data[i]['PVT_region'])['compr']
                    data_from_PVTO.append(info_from_PVTO)
                PVT_data=data_from_PVTO
    
        PVDG_tables=pebi.get_number_of_pvt_regions(model_id)['PVDG'] #число таблиц PVDG
        if PVDG_tables!=0:
            data_from_PVDG=[] #информация по газу
            for i in range(0,obj_number):
                info_from_PVDG={}
                info_from_PVDG['obj_name']=data[i]['obj_name'] #имя объекта
                info_from_PVDG['B_gas']=pebi.get_pvdg(model_id,data[i]['PVT_region'])['table'][1] #об. коэф-т расширения газа
                info_from_PVDG['gas_visc']=pebi.get_pvdg(model_id,data[i]['PVT_region'])['table'][2] #вязкость газа
                data_from_PVDG.append(info_from_PVDG)
        
        data_from_PVTW=[] #информация по воде
        for i in range(0,obj_number):
            info_from_PVTW={}
            info_from_PVTW['obj_name']=data[i]['obj_name'] #имя объекта
            info_from_PVTW['B_water']=pebi.get_pvtw(model_id,data[i]['PVT_region'])['table'][1] #коэф-т об. расширения воды
            info_from_PVTW['water_comp']=pebi.get_pvtw(model_id,data[i]['PVT_region'])['table'][2] #сжимаеомость воды
            info_from_PVTW['water_visc']=pebi.get_pvtw(model_id,data[i]['PVT_region'])['table'][3] #вязкость воды
            data_from_PVTW.append(info_from_PVTW)
        #print(data_from_PVTO)      #дебажный принт
        
        PVTO_tables=pebi.get_number_of_pvt_regions(model_id)['PVTO'] #число таблиц PVTO
        
        tables=(PVDO_tables,PVCDO_tables,PVTO_tables) #сейчас будем считать число регионов PVT
        for el in tables:
            if el!=0:
                number_of_PVT_regions=el #количество регионов PVT
        
        RPP=0 #список с значениями числа таблиц ОФП для каждого региона
        for k in range(0,len(data)):
            RPP+=pebi.get_number_of_relperm_tables(model_id, data[i]['PVT_region'])['SWOF'] #идентификация по числу таблиц ОФП вода-нефть
        
        dens_data=[] #ифнормация по плотностям
        for i in range(0,obj_number):
            info_from_dens={}
            info_from_dens['obj_name']=data[i]['obj_name'] #имя объекта
            info_from_dens['oil_dens']=pebi.get_densities(model_id, data[i]['PVT_region'])['oil']
            info_from_dens['gas_dens']=pebi.get_densities(model_id, data[i]['PVT_region'])['gas']
            info_from_dens['water_dens']=pebi.get_densities(model_id, data[i]['PVT_region'])['water']
            dens_data.append(info_from_dens)
        
        dims=pebi.get_dimens(model_id) #размерность сетки
        n_cells=dims['nx']*dims['ny']*dims['nz'] #число ячеек
        pebi.init(model_id) #проинициализировали модель, чтобы забрать кубы
        
        NTG=pebi.get_ntg(model_id)['values'] #список значений NTG для каждой ячейки
        PERMX=pebi.get_permx(model_id)['values'] #список значений PERMX для каждой ячейки
        PERMY=pebi.get_permy(model_id)['values'] #список значений PERMY для каждой ячейки
        PERMZ=pebi.get_permz(model_id)['values'] #список значений PERMZ для каждой ячейки
        PORO=pebi.get_poro(model_id)['values'] #список значений PORO для каждой ячейки
        ACTNUM=pebi.get_actnum(model_id)['values'] #список значений ACTNUM для каждой ячейки
        FIPNUM=pebi.get_fipnum(model_id)['values'] #список значений FIPNUM для каждой ячейки
        SWAT=pebi.get_swat(model_id)['values'] #список значений Sw для кажждой ячейки
        SGAS=pebi.get_sgas(model_id)['values'] #список значений Sg для каждой ячейки
        
        DEPTH=pebi.get_depth(model_id)['values'] #куб глубин
        
        SOIL=[]
        for el in range(0,len(SWAT)):
            SOIL.append(1-SWAT[el])
        
        n_active_cells=0 #изначальное число активных ячеек
        for el in ACTNUM:
            if el==1.0:
                n_active_cells+=1
       
        obj_NTG=get_arr_layer(obj_number, data, dims, NTG) #сформировали список словарей NTG для каждого объекта
        
        obj_SWAT=get_arr_layer(obj_number,data,dims,SWAT) #сформировали список словарей Sw для каждого объекта
        obj_SOIL=get_arr_layer(obj_number,data,dims,SOIL) #сформировали список словарей So для каждого объекта
        
        obj_PERMX=get_arr_layer(obj_number, data, dims, PERMX) #сформировали список словарей PERMX для каждого объекта
        obj_PERMY=get_arr_layer(obj_number, data, dims, PERMY) #сформировали список словарей PERMY для каждого объекта
        obj_PERMZ=get_arr_layer(obj_number, data, dims, PERMZ) #сформировали список словарей PERMZ для каждого объекта
        
        obj_PORO=get_arr_layer(obj_number, data, dims, PORO) #сформировали список словарей PERMZ для каждого объекта
        
        number_of_wells=len(pebi.get_well_list(model_id)) #число скважин в модели 
        
        
        download_index=50 #выполнили половину - скомпоновали данные
        self.ui.progressBar.setValue(download_index)
        
        file_out=xlsxwriter.Workbook('Паспорт ГД модели.xlsx')
        sheet_1=file_out.add_worksheet('ГФХ')
        format_1=file_out.add_format({'bold':True,'align':'center'}) #формат заголовков
        format_2=file_out.add_format({'bold':False,'align':'center'}) #формат незаголовков
        #заполняем шапку
        sheet_1.set_column(0,0, 60)
        sheet_1.write(1,0,'Параметр',format_1)
        for n in range(0,obj_number):
            sheet_1.write(1,n+1,data[n]['obj_name'],format_1) #имена объектов
        GFH=['Средняя глубина залегания кровли, м','Абсолютная отметка ВНК, м','Абсолютная отметка ГНК, м',
             'Опорная глубина, м','Пластовое давление','Объёмный коэффициент нефти, д. ед.','Давление насыщения, бар','Газосодержание, м3/т',
             'Вязкость нефти в пл. условиях, сПз','Сжимаемость породы, 1/бар',
             'Проницаемость, мД','Пористость, д. ед.','Песчанистость, д. ед.','Нефтенасыщенность, д. ед.',
             'Плотность нефти в поверхностных условиях породы, кг/м3','Плотность воды в поверхностных условиях породы, кг/м3',
             'Средняя общая толщина, м','Средняя эффективная толщина, м']
        for n in range(0,len(GFH)):
            sheet_1.write(n+2,0,GFH[n],format_2)
        #заполняем значениями
        for n in range(0,obj_number):
            sheet_1.write(2,n+1,average_layer_param(n, data, data[n]['top'], dims, DEPTH)['average'])
            sheet_1.write(3,n+1,data_from_equil[n]['hwoc'],format_2) #ВНК
            sheet_1.write(4,n+1,data_from_equil[n]['hgoc'],format_2) #ГНК
            sheet_1.write(5,n+1,data_from_equil[n]['href'],format_2) #опорная глубина
            sheet_1.write(6,n+1,data_from_equil[n]['ref_pres'],format_2) #пластовое давление
            sheet_1.write(12,n+1,(obj_PERMX[n]['average']+obj_PERMY[n]['average']+obj_PERMZ[n]['average'])/3,format_2) #средняя проницаемость
            sheet_1.write(13,n+1,obj_PORO[n]['average'],format_2) #средняя пористость
            sheet_1.write(14,n+1,obj_NTG[n]['average'],format_2) #средняя песчанистость
            sheet_1.write(15,n+1,obj_SOIL[n]['average'],format_2) #средняя нефтенасыщенность
            sheet_1.write(16,n+1,dens_data[n]['oil_dens'],format_2) #плотность нефти
            sheet_1.write(17,n+1,dens_data[n]['water_dens'],format_2) #плотность воды
            sheet_1.write(18,n+1,(avg_thick(obj_number, data, dims, DEPTH))[n]['avg_thickness'],format_2) #средняя общая толщина
            sheet_1.write(19,n+1,(avg_thick(obj_number, data, dims, DEPTH))[n]['avg_thickness']*obj_NTG[n]['average'],format_2) #ср. эфф. толщина
            #с PVTями посложнее
            if (PVDO_tables==0):
                if (PVCDO_tables==0): #если у нас табличка PVTO
                    sheet_1.write(7,n+1,'-',format_2) #объемный коэффициент нефти
                    sheet_1.write(8,n+1,data_from_PVTO[n]['bub_pres'],format_2) #давление насыщения
                    sheet_1.write(9,n+1,data_from_PVTO[n]['Rs'],format_2) #газосодержание
                    sheet_1.write(10,n+1,data_from_PVTO[n]['visc'],format_2) #вязкость нефти в пл. условиях
                    sheet_1.write(11,n+1,data_from_PVTO[n]['rock'],format_2) #сжимаемость породы
                elif (PVTO_tables==0): #если у нас табличка PVCDO
                    sheet_1.write(7,n+1,data_from_PVCDO[n]['B_oil'],format_2) #объёмник
                    sheet_1.write(8,n+1,'-',format_2) #давление насыщения
                    sheet_1.write(9,n+1,'-',format_2) #газосодержание
                    sheet_1.write(10,n+1,data_from_PVCDO[n]['visc'],format_2) #вязкость в пл. условиях
                    sheet_1.write(11,n+1,data_from_PVCDO[n]['rock'],format_2) #сжимаемость породы
            else: #если у нас табличка PVDO
                pass #пока что pass
            download_index+=25/obj_number
            self.ui.progressBar.setValue(int(math.trunc(download_index)))
        
        sheet_2=file_out.add_worksheet('Описание модели')
        #шапка
        sheet_2.set_column(0,1, 40)
        
        sheet_2.write(1,0,'Параметр',format_1)
        sheet_2.write(1,1,'Модель',format_1)
        MODEL_PARAMS=['Размерность сетки','Наличие/отсутствие апскейлинга','Наличие/отсутствие разломов','Количество ячеек (corner point)',
                      'Количество активных ячеек (corner point)','Количество ячеек (PEBI)','Количество активных ячеек (PEBI)',
                      'Количество скважин','Количество лет разработки','Количество регионов подсчёта запасов', 'Количество регионов инициализации',
                      'Количество регионов PVT','Наличие аквиферов','Наличие множителей порового объема','Модель PVT свойств','Число регионов ОФП']
        for n in range(0,len(MODEL_PARAMS)):
            sheet_2.write(n+2,0,MODEL_PARAMS[n],format_2)
        #заполняем значениями
        sheet_2.write(2,1,f"{dims['nx']}x{dims['ny']}x{dims['nz']}",format_2) #размерность сетки
        sheet_2.write(5,1,n_cells,format_2) #число ячеек (геометрия угловой точки)
        sheet_2.write(6,1,n_active_cells,format_2) #активных (геометрия угловой точки)
        sheet_2.write(9,1,number_of_wells,format_2) #число скважин
        sheet_2.write(11,1,unique_num(FIPNUM),format_2) #количество регионов подсчета запасов
        sheet_2.write(12,1,EQ_number,format_2) #число регионов инициализации
        sheet_2.write(13,1,number_of_PVT_regions,format_2) #число регионов PVT
        sheet_2.write(16,1,model_type(PVTO_tables),format_2)
        sheet_2.write(17,1,RPP,format_2) #число регионов ОФП
        download_index+=25
        dict_PORO=bar_graph_data(PORO) #значения для гистограммы пористости
        dict_PERMX=bar_graph_data(PERMX) #значения для гистограммы проницаемости по Х
        dict_PERMY=bar_graph_data(PERMY) #значения для гистограммы проницаемости по Y
        dict_PERMZ=bar_graph_data(PERMZ) #значения для гистограммы проницаемости по Z
        dict_NTG=bar_graph_data(NTG) #значения для гистограммы песчанистости
        
        #сопоставление с геологической моделью
        sheet_3=file_out.add_worksheet('Сопоставление ГМ и ГДМ')
        #шапка
        sheet_3.set_column(0,1,35)
        sheet_3.write(1,0,'Параметр',format_1)
        sheet_3.write(1,1,'Единицы измерения',format_1)
        sheet_3.write(1,2,'ГМ',format_1)
        sheet_3.set_column(3,3,25)
        sheet_3.write(1,3,'Отклонение (%)',format_1)
        
        ROWS=['Начальные геологические запасы нефти','Поровый объем','Объём нефтенасыщенных пород','Ср. эфф. насыщенная толщина','Ср. коэф-т пористости',
                 'Ср. коэф-т проницаемости']
        METRICS=['тыс. т','тыс. м3','тыс. м3','тыс. м2','м','д. ед.','д. ед.']
        stroka=2
        for n in range(0,obj_number):
            sheet_3.merge_range(stroka,0,stroka,3,data[n]['obj_name'],format_1)
            sheet_3.write(stroka+1,0,ROWS[0],format_2)
            sheet_3.write(stroka+1,1,METRICS[0],format_2)
            sheet_3.write(stroka+2,0,ROWS[1],format_2)
            sheet_3.write(stroka+2,1,METRICS[1],format_2)
            sheet_3.write(stroka+3,0,ROWS[2],format_2)
            sheet_3.write(stroka+3,1,METRICS[2],format_2)
            sheet_3.write(stroka+4,0,ROWS[3],format_2)
            sheet_3.write(stroka+4,1,METRICS[3],format_2)
            sheet_3.write(stroka+5,0,ROWS[4],format_2)
            sheet_3.write(stroka+5,1,METRICS[4],format_2)
            sheet_3.write(stroka+6,0,ROWS[5],format_2)
            sheet_3.write(stroka+6,1,METRICS[5],format_2)
            sheet_3.write(stroka+7,0,ROWS[6],format_2)
            sheet_3.write(stroka+7,1,METRICS[6],format_2)
            stroka+=8
        self.ui.progressBar.setValue(int(math.trunc(download_index)))
        file_out.close()
        debug=open('debug.txt','w')
        debug_keys=open('debug_keys.txt','w')
        debug.write(str(dict_PORO))
        for key in pebi.get_events(model_id):
            debug_keys.write(str(key)+'\n')
        debug.close()
        debug_keys.close()
        os.chdir(current_folder) #вышли обратно в папку с скриптом
        

import sys
app = QtWidgets.QApplication([])
application = mywindow()
application.show()
sys.exit(app.exec())