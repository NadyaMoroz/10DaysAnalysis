import numpy as np
import pandas as pd
import math
import datetime
from tabulate import tabulate

#импортируем файл и подгружаем в датафрейм
def import_file():
    while True:
        try:
            path = input('Путь к файлу, включая его название: ').strip().replace('\\','/')
            sheet = input('Название листа в файле: ').strip()
            matrix = pd.read_excel(path,sheet_name=sheet)
            break
        except:
            print('Не могу найти такой файл и вкладку :(')
    return matrix
#предлагаем пользователю список колонок и предлагаем выбрать одну
def list_ids(default_list):
    print()
    list_length = len(default_list)
    list_index = [i for i in range(len(default_list))]
    for i in range(list_length-1,0,-1):
        if list_length % i == 0:
            n1 = 0
            n2 = i
            break
    while n2 <= list_length:
        if list_length // n2 == 0:
            n2 = list_length
        df = pd.DataFrame(columns=default_list[n1:n2])
        df.loc[''] = list_index[n1:n2]
        print()
        print(tabulate(df, headers='keys', tablefmt='grid',showindex=False))
        n1 += i
        n2 += i
#ввод пользователем индекса колонки
def user_column_input(default_list):
    while True:
        user_answer = input('Введите номер столбца, который хотите использовать')
        if user_answer.isdigit():
            if int(user_answer) < len(default_list):
                return int(user_answer)
#сопоставляем нужные данные с индексами колонок
def choose_columns(searched_data, default_df):
    list_ids(default_df.columns)
    columns_list = []
    new_names = []
    for data_name in searched_data:
        print(data_name + ': ')
        column_number = user_column_input(default_df.columns)
        columns_list.append(int(column_number))
        new_names.append(data_name)
    final_data = default_df[[columns_list]]
    final_data.columns = new_names
    return final_data

print('Загрузи файл, в котором содержится информация о заказах: дата, сумма, идентификатор заказа, идентификатор пользователя')
orders = import_file()
print('Вот список доступных колонок. Сопоставь их номера с теми данными, которые нужны для расчета: ')
cleared_orders = choose_columns(['Дата','Сумма', 'Идентификатор заказа', 'Идентификатор пользователя'], orders)
print(tabulate(cleared_orders.head(5), headers='keys', tablefmt='grid',showindex=False))

print('Загрузи файл, в котором содержится информация об открытии рассылок: идентификатор пользователя, дата открытия, название рассылки')
openings = import_file()




commonStatistics = pd.DataFrame(columns=['month','Общая выручка', 'Количество заказов', 'Общая выручка email 10', 'Общее количество email 10', 
                                        'Общая выручка ga 10', 'Общее количество ga 10', 'Выручка с триггеров 10',
                                         'Выручка с регулярки 10', 'Общая выручка email 5', 'Общее количество email 5',
                                        'Общая выручка ga 5', 'Общее количество ga 5', 'Выручка с триггеров 5',
                                         'Выручка с регулярки 5'])
statisticnum = 0
import datetime
for x in range(len(startdate)):
    print('Month: ', startdate[x])
    ordersPay = ordersPayAll[(ordersPayAll['Дата заказа'] >= startdate[x])&(ordersPayAll['Дата заказа'] <= enddate[x])]
    ordersPay['Sales'] = ordersPay['Sales'].astype('str').str.replace(' ','').astype('float')
    opensAll = opensAll1[(opensAll1['Дата открытия'] >= startdate1[x])&(opensAll1['Дата открытия'] <= enddate[x])]
    #приводим почты к стандартному виду (низкий регистр)
    opensAll['Email1'] = opensAll['Email'].str.lower()
    ordersPay['Email1'] = ordersPay['Email'].str.lower()
    #добавляем столбец с только датой открытия (без времени)
    opensAll['Дата открытия1'] = opensAll['Дата открытия'].dt.date
    #получаем список почт, которые открыли письмо и что-то купили
    emails1 = ordersPay.merge(opensAll, left_on = 'Email1', right_on = 'Email1')
    emailList = emails1[['Email1']].groupby(['Email1']).nunique()
    #ищем для каждого пользователя все заказы, для каждого заказа прикрепляем ближайшее письмо (в пределах daysNumber, если такое есть). 
    #Потом группируем по письмам для каждого пользователя и добавляем в общую таблицу. Так мы получаем статистику по юзерам — кто после каких писем купил и сколько
    ordersGrouped5 = pd.DataFrame(columns=['ID клиента','Email', 'Сумма', 'Название рассылки','Количество заказов'])
    ordersGA5 = pd.DataFrame(columns=['ID клиента','Email', 'Сумма', 'Название рассылки','ID заказа'])
    ordersGrouped10 = pd.DataFrame(columns=['ID клиента','Email', 'Сумма', 'Название рассылки','Количество заказов'])
    ordersGA10 = pd.DataFrame(columns=['ID клиента','Email', 'Сумма', 'Название рассылки','ID заказа'])
    grouped = [ordersGrouped5, ordersGrouped10]
    ordersga = [ordersGA5, ordersGA10]
    daysNumber = [2,4]
    for iteration in range(2):
        print('iteration days: ', daysNumber[iteration])
        print('')
        n = 0
        n1 = 0
        for item in emailList.index.get_level_values(level=0):
            result = ordersPay[ordersPay['Email1'] == item].sort_values('Дата заказа')
            ordersGrouped1 = pd.DataFrame(columns=['ID клиента','Email', 'Сумма', 'Название рассылки','ID заказа'])
            for orders in range(result.shape[0]):
                for i in range(daysNumber[iteration]):
                    letterDate = (datetime.datetime.strptime(str(result.iloc[orders,7]), '%Y-%m-%d %H:%M:%S') - datetime.timedelta(days=i)).date()
                    opensWork = opensAll[(opensAll['Email1'] == item)&(opensAll['Дата открытия1'] == letterDate)]
                    if opensWork.shape[0] > 0:
                        opensWork.sort_values('Дата открытия',ascending=False,inplace=True)
                        if str(opensWork.iloc[0,2]) == 'Welcome 1.1 Nikon Store Открытие':
                            if i != 0:
                                ordersGrouped1.loc[orders] = opensWork.iloc[0,0], opensWork.iloc[0,1], result.iloc[orders,4], opensWork.iloc[0,2], result.iloc[orders,0]
                                ordersga[iteration].loc[n1] = opensWork.iloc[0,0], opensWork.iloc[0,1], result.iloc[orders,4], opensWork.iloc[0,2], result.iloc[orders,0]
                                n1 += 1
                        else: 
                            ordersGrouped1.loc[orders] = opensWork.iloc[0,0], opensWork.iloc[0,1], result.iloc[orders,4], opensWork.iloc[0,2], result.iloc[orders,0]
                            ordersga[iteration].loc[n1] = opensWork.iloc[0,0], opensWork.iloc[0,1], result.iloc[orders,4], opensWork.iloc[0,2], result.iloc[orders,0]
                            n1 += 1
                        break
            ordersGrouped1 = ordersGrouped1.groupby(['Название рассылки','ID клиента','Email']).agg({'ID заказа':'count','Сумма':'sum'})
            ordersGrouped1.reset_index(inplace=True)
            for i in range(ordersGrouped1.shape[0]):
                grouped[iteration].loc[n] = ordersGrouped1.iloc[i,1],ordersGrouped1.iloc[i,2],ordersGrouped1.iloc[i,4],ordersGrouped1.iloc[i,0],ordersGrouped1.iloc[i,3]
                n += 1
        print('Email generation finished.')
        grouped[iteration]['Тип'] = np.where((grouped[iteration]['Название рассылки'].str.contains('Триггер') == True)|
                                (grouped[iteration]['Название рассылки'].str.contains('Welcome') == True)|
                                (grouped[iteration]['Название рассылки'].str.contains('ЕА') == True)|
                                (grouped[iteration]['Название рассылки'].str.contains('Магистратура') == True)|
                                (grouped[iteration]['Название рассылки'].str.contains('Дозаполнение анкеты') == True),
                                'Триггер','Регулярная рассылка')
        print('Common data finished.')
    
    GACommon10 = ordersGA10.merge(GA, left_on='ID заказа', right_on = 'Идентификатор транзакции', how='left')
    GACommon10['Type'] = np.where(GACommon10['Type'].isnull() == True, 'other', GACommon10['Type'] )
    GACommon10['Сумма'] = np.where(GACommon10['Type'] == 'cpc', 
                                   GACommon10['Сумма'].astype('str').str.replace(' ','').astype('float')/2, 
                                   GACommon10['Сумма'].astype('str').str.replace(' ','').astype('float'))
    GACommon10 = GACommon10[['ID клиента','Email','Сумма','Название рассылки','ID заказа','Type']]
    
    GACommon5 = ordersGA5.merge(GA, left_on='ID заказа', right_on = 'Идентификатор транзакции', how='left')
    GACommon5['Type'] = np.where(GACommon5['Type'].isnull() == True, 'other', GACommon5['Type'] )
    GACommon5['Сумма'] = np.where(GACommon5['Type'] == 'cpc',
                                  GACommon5['Сумма'].astype('str').str.replace(' ','').astype('float')/2, 
                                  GACommon5['Сумма'].astype('str').str.replace(' ','').astype('float'))
    GACommon5 = GACommon5[['ID клиента','Email','Сумма','Название рассылки','ID заказа','Type']]
    
    ordersGrouped10['Сумма'] = ordersGrouped10['Сумма'].astype('str').str.replace(' ','').astype('float')
    ordersGrouped5['Сумма'] = ordersGrouped5['Сумма'].astype('str').str.replace(' ','').astype('float')
    
    ordersGrouped5Common = ordersGrouped5[['Название рассылки','Количество заказов','Сумма']].groupby(['Название рассылки']).agg({'Сумма':'sum','Количество заказов':'sum'})
    ordersGrouped5Common.loc['Итог'] = ordersGrouped5Common['Сумма'].sum(), ordersGrouped5Common['Количество заказов'].sum()
    ordersGrouped5Common.reset_index(inplace=True)
    ordersGrouped5Common.rename(columns={'Сумма':'Сумма 1', 'Количество заказов':'Количество заказов 1'}, inplace=True)

    ordersGrouped5Trig = ordersGrouped5[ordersGrouped5['Тип'] == 'Триггер']
    ordersGrouped5Reg = ordersGrouped5[ordersGrouped5['Тип'] != 'Триггер']

    ordersGrouped5Trig = ordersGrouped5Trig[['Название рассылки','Количество заказов','Сумма']].groupby(['Название рассылки']).agg({'Сумма':'sum','Количество заказов':'sum'})
    ordersGrouped5Trig.loc['Итог'] = ordersGrouped5Trig['Сумма'].sum(), ordersGrouped5Trig['Количество заказов'].sum()
    ordersGrouped5Trig.reset_index(inplace=True)
    ordersGrouped5Trig.rename(columns={'Сумма':'Сумма 1', 'Количество заказов':'Количество заказов 1'}, inplace=True)

    ordersGrouped5Reg = ordersGrouped5Reg[['Название рассылки','Количество заказов','Сумма']].groupby(['Название рассылки']).agg({'Сумма':'sum','Количество заказов':'sum'})
    ordersGrouped5Reg.loc['Итог'] = ordersGrouped5Reg['Сумма'].sum(), ordersGrouped5Reg['Количество заказов'].sum()
    ordersGrouped5Reg.reset_index(inplace=True)
    ordersGrouped5Reg.rename(columns={'Сумма':'Сумма 1', 'Количество заказов':'Количество заказов 1'}, inplace=True)

    ordersGA5 = GACommon5[['Название рассылки','ID заказа','Сумма']].groupby(['Название рассылки']).agg({'Сумма':'sum','ID заказа':'count'})
    ordersGA5.loc['Итог'] = ordersGA5['Сумма'].sum(), ordersGA5['ID заказа'].sum()
    ordersGA5.reset_index(inplace=True)
    ordersGA5.rename(columns={'Сумма':'Сумма 1','ID заказа':'Количество заказов 1'},inplace=True)
    
    ordersGrouped10Common = ordersGrouped10[['Название рассылки','Количество заказов','Сумма']].groupby(['Название рассылки']).agg({'Сумма':'sum','Количество заказов':'sum'})
    ordersGrouped10Common.loc['Итог'] = ordersGrouped10Common['Сумма'].sum(), ordersGrouped10Common['Количество заказов'].sum()
    ordersGrouped10Common.reset_index(inplace=True)
    ordersGrouped10Common.rename(columns={'Сумма':'Сумма 3', 'Количество заказов':'Количество заказов 3'}, inplace=True)

    ordersGrouped10Trig = ordersGrouped10[ordersGrouped10['Тип'] == 'Триггер']
    ordersGrouped10Reg = ordersGrouped10[ordersGrouped10['Тип'] != 'Триггер']

    ordersGrouped10Trig = ordersGrouped10Trig[['Название рассылки','Количество заказов','Сумма']].groupby(['Название рассылки']).agg({'Сумма':'sum','Количество заказов':'sum'})
    ordersGrouped10Trig.loc['Итог'] = ordersGrouped10Trig['Сумма'].sum(), ordersGrouped10Trig['Количество заказов'].sum()
    ordersGrouped10Trig.reset_index(inplace=True)
    ordersGrouped10Trig.rename(columns={'Сумма':'Сумма 3', 'Количество заказов':'Количество заказов 3'}, inplace=True)

    ordersGrouped10Reg = ordersGrouped10Reg[['Название рассылки','Количество заказов','Сумма']].groupby(['Название рассылки']).agg({'Сумма':'sum','Количество заказов':'sum'})
    ordersGrouped10Reg.loc['Итог'] = ordersGrouped10Reg['Сумма'].sum(), ordersGrouped10Reg['Количество заказов'].sum()
    ordersGrouped10Reg.reset_index(inplace=True)
    ordersGrouped10Reg.rename(columns={'Сумма':'Сумма 3', 'Количество заказов':'Количество заказов 3'}, inplace=True)

    ordersGA10 = GACommon10[['Название рассылки','ID заказа','Сумма']].groupby(['Название рассылки']).agg({'Сумма':'sum','ID заказа':'count'})
    ordersGA10.loc['Итог'] = ordersGA10['Сумма'].sum(), ordersGA10['ID заказа'].sum()
    ordersGA10.reset_index(inplace=True)
    ordersGA10.rename(columns={'Сумма':'Сумма 3','ID заказа':'Количество заказов 3'},inplace=True)

    ordersReg = ordersGrouped10Reg.merge(ordersGrouped5Reg, how='left')
    ordersTrig = ordersGrouped10Trig.merge(ordersGrouped5Trig, how='left')
    ordersCommon = ordersGrouped10Common.merge(ordersGrouped5Common, how='left')
    ordersGA = ordersGA10.merge(ordersGA5, how='left')
    
    commonStatistics.loc[statisticnum] = startdate[x][:7], ordersPay['Sales'].sum(),ordersPay['ID'].count(), ordersCommon.iloc[ordersCommon.shape[0]-1,1], ordersCommon.iloc[ordersCommon.shape[0]-1,2], ordersGA.iloc[ordersGA.shape[0]-1,1], ordersGA.iloc[ordersGA.shape[0]-1,2], ordersTrig.iloc[ordersTrig.shape[0]-1,1], ordersReg.iloc[ordersReg.shape[0]-1,1], ordersCommon.iloc[ordersCommon.shape[0]-1,3], ordersCommon.iloc[ordersCommon.shape[0]-1,4],ordersGA.iloc[ordersGA.shape[0]-1,3],ordersGA.iloc[ordersGA.shape[0]-1,4], ordersTrig.iloc[ordersTrig.shape[0]-1,3], ordersReg.iloc[ordersReg.shape[0]-1,3]
    statisticnum += 1
    with pd.ExcelWriter('Nikon-' + startdate[x][:7] + '.xlsx') as writer:  
        ordersCommon.to_excel(writer, sheet_name='Сводная общая')
        ordersTrig.to_excel(writer, sheet_name='Сводная триггер')
        ordersReg.to_excel(writer, sheet_name='Сводная регулярные')
        ordersGA.to_excel(writer, sheet_name='Сводная GA')
        ordersGrouped5.to_excel(writer, sheet_name='Исходные данные 5 дней')
        ordersGrouped10.to_excel(writer, sheet_name='Исходные данные 10 дней')
        writer.save()


#период
#сколько дней прошло с рассылки

