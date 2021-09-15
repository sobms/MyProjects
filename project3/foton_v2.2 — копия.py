import pandas as pd
import numpy as np
import foton_display4
from PyQt5 import QtWidgets
import sys
import matplotlib.ticker as ticker
from matplotlib.backends.backend_qt4agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from scipy.interpolate import make_interp_spline

class App(QtWidgets.QMainWindow, foton_display4.Ui_MainWindow):
    

    def __init__(self):
        self.a1 = 0
        self.b1 = 0
        self.a2 = 0
        self.b2 = 0
        self.S = 0
        self.N_max = 0
        self.Nu_i = 0
        self.P = 0
        self.E_ey = 0
        self.E_ey_nom = 55
        self.P_max = 0
        self.table = 0
        self.N_min = 0
        self.N_max_ey = 0
        self.N1_te = 0
        self.E_max_h2 = 0
        self.E_vp = 0
        self.flag = False
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.load_table)
        self.pushButton_6.clicked.connect(self.Count_N_min)
        self.pushButton_2.clicked.connect(self.Calc_ey_efficiency)
        self.pushButton_3.clicked.connect(self.Calc_ey_num)
        self.pushButton_4.clicked.connect(self.Calc_te_num)
        self.pushButton_5.clicked.connect(self.Calc_num_bal_and_max_mass)
        self.pushButton_7.clicked.connect(self.Calc_stoves)
        self.pushButton_8.clicked.connect(self.Calc_boiler)




    def load_table(self):
        filename, filetype = QtWidgets.QFileDialog.getOpenFileName(self,
                "Выбрать файл",
                ".",
                "Excel Files(*.xlsx);;")
        self.table = pd.read_excel(filename, usecols=['Год','Месяц','День','Час','Энергия_излуч.','Нагрузка'])
        Max_Row = self.table.shape[0]
        self.E_ey = np.zeros(Max_Row)
        self.P = np.zeros(Max_Row)
        self.P_max = np.max(self.table['Нагрузка'])        
        #p_nom = 4
        self.tableWidget.setRowCount(Max_Row)
        for row in range(0,Max_Row):
            for column in range(0,6):
                item = QtWidgets.QTableWidgetItem()
                self.tableWidget.setItem(row, column, item)
                item.setText(str(self.table.iloc[row,column]))

    def renew_table(self):
        Max_Row = self.table.shape[0]
        for row in range(0,Max_Row):
            for column in range(0,6):
                item = QtWidgets.QTableWidgetItem()
                self.tableWidget.setItem(row, column, item)
                item.setText(str(self.table.iloc[row,column]))
                de_plus = []
        de_minus = []
        for i in range(0,Max_Row):
                self.E_ey[i] = self.table['Энергия_излуч.'].iloc[i]*self.S/1000000
                self.P[i] = self.table['Нагрузка'].iloc[i]
                dE = (self.E_ey[i]*self.N_min-self.P[i])
                if dE>0:
                    de_plus.append(np.round_((dE*self.Nu_i*(self.a1*(self.E_ey_nom-self.E_ey[i]*self.N_min)+self.b1)),4))
                    de_minus.append(0)
                elif dE<=0:
                    de_minus.append(np.round_((-dE/(self.Nu_i*(self.a2*(np.round_(self.P_max)-self.P[i])+self.b2))),4))
                    de_plus.append(0)
        self.table['+dE'] =de_plus
        self.table['-dE'] =de_minus

    def CalcD_n(self, table, n):
        F1 = 0
        F2 = 0
        Max_Row = table.shape[0]
        P_max = np.max(table['Нагрузка'])
        for i in range(0,Max_Row):
            self.E_ey[i] = (table['Энергия_излуч.'].iloc[i]+table['Энергия_излуч.'].iloc[i-1])/2*self.S/1000000
            self.P[i] = (table['Нагрузка'].iloc[i]+table['Нагрузка'].iloc[i-1])/2
            dE = self.E_ey[i]*n-self.P[i]*self.Nu_i
            if dE>0:
                F1+= dE*(self.a1*(dE-self.E_ey_nom) + self.b1)
            if dE<=0:
                F2+= (-dE-self.E_vp)/(self.a2*(P_max-dE) + self.b2)+self.E_vp
        D_n = np.abs(F1-F2)
        return D_n  

    def most_favorable_day(self):
        data=self.table.drop(columns=['Час']).groupby(by = ['Год','Месяц','День']).sum()
        return np.array(data.loc[data['Энергия_излуч.']==np.max(data['Энергия_излуч.'])].index)[0]
    
    def get_daily_power(self, month, day):
        data = self.table.loc[(self.table['День'] == day) & (self.table['Месяц'] == month)]
        y1 = []
        y2 = []
        sum_ = 0
        for row in range(0,data.shape[0]):
            y1.append(data.iloc[row]['Энергия_излуч.']*self.S/1000000*self.N_min)
            y2.append(data.iloc[row]['Нагрузка'])
            sum_ +=data.iloc[row]['Энергия_излуч.']*self.S/1000000*self.N_min
        x = [(i+1) for i in range(0, data['Энергия_излуч.'].shape[0])]
        fig = Figure(figsize=(6,2))
        fig.patch.set_facecolor('#b6c3be')
        ax = fig.add_subplot()
        ax.set_title('Самый благоприятный день по приходу солнечной энергии ' + str(day) + ' число ' + str(month) + ' мес', fontsize = 8, horizontalalignment='center')
        ax.set_xlabel('Время суток, час', fontsize=4)
        ax.set_ylabel('Мощность, кВт', fontsize=8)
        ax.grid()
        ax.patch.set_facecolor('#e6e6e6')
        ax.tick_params(axis = 'both', which = 'major', labelsize = 6)
        ax.xaxis.set_major_locator(ticker.MultipleLocator(1))
        ax.annotate(str(np.round_(sum_))+" кВт*ч", xy=(0, (np.max(y1)-np.min(y1))/10*9), fontsize = 6)
        ax.plot(x,y1, color='orange')
        ax.plot(x,y2,color='green')
        canvas = FigureCanvas(fig)
        scene = QtWidgets.QGraphicsScene(self.graphicsView_2)
        scene.addWidget(canvas)
        self.graphicsView_2.setScene(scene)

    def get_daily_workload(self, month, day):
        data = self.table.loc[(self.table['День'] == day) & (self.table['Месяц'] == month)]
        y = []
        sum_=0
        for wkload in data['Нагрузка']:
            sum_+=wkload
            y.append(wkload)
        x = [(i+1) for i in range(0, data['Нагрузка'].shape[0])]
        fig = Figure(figsize=(6,2))
        ax = fig.add_subplot()
        fig.patch.set_facecolor('#b6c3be')
        ax.set_title('Суточный график нагрузки ' + str(day) + ' число ' + str(month) + ' мес', fontsize = 8, horizontalalignment='center')
        ax.set_xlabel('Время суток, ч', fontsize=5)
        ax.set_ylabel('Нагрузка, кВт', fontsize=8)
        ax.annotate(str(np.round_(sum_))+" кВт*ч", xy=(0, (np.max(y)-np.min(y))/10*9), fontsize = 6)
        ax.grid()
        ax.patch.set_facecolor('#e6e6e6')
        ax.tick_params(axis = 'both', which = 'major', labelsize = 6)
        ax.xaxis.set_major_locator(ticker.MultipleLocator(1))
        spl = make_interp_spline(x, y, k=2)
        xnew = np.linspace(np.min(x), np.max(x), 200)
        ysmooth = spl(xnew)
        ax.plot(xnew,ysmooth, color = 'green')
        canvas = FigureCanvas(fig)
        scene = QtWidgets.QGraphicsScene(self.graphicsView_3)
        scene.addWidget(canvas)
        self.graphicsView_3.setScene(scene)

    def get_year_dE(self):
        y = []
        x = ['янв','фев','мар','апр','май','июн','июл','авг','сен','окт','ноя','дек']
        P_max = np.max(self.table['Нагрузка'])
        for i in range(0,12):
            data = self.table.loc[self.table['Месяц']==(i+1)]
            F1 = 0
            F2 = 0
            for i in range(0,data.shape[0]):
                self.E_ey[i] = (data['Энергия_излуч.'].iloc[i]+data['Энергия_излуч.'].iloc[i-1])/2*self.S/1000000
                self.P[i] = (data['Нагрузка'].iloc[i]+data['Нагрузка'].iloc[i-1])/2
                dE = self.E_ey[i]*self.N_min-self.P[i]*self.Nu_i
                if dE>0:
                    F1+= dE*(self.a1*(dE-self.E_ey_nom) + self.b1)
                if dE<=0:
                    F2+= (-dE-self.E_vp)/(self.a2*(P_max-dE) + self.b2)+self.E_vp
            y.append(F1-F2)
        fig = Figure(figsize=(6,2))
        fig.patch.set_facecolor('#b6c3be')
        ax = fig.add_subplot()
        ax.set_title('Годовой баланс энергии запасенной в водороде', fontsize = 8, horizontalalignment='center')
        ax.set_xlabel('Месяц', fontsize=3)
        ax.set_ylabel('dE, кВт', fontsize=8)
        ax.grid()
        ax.tick_params(axis = 'both', which = 'major', labelsize = 6)
        ax.plot(x,np.zeros(12),color = 'black')
        ax.patch.set_facecolor('#e6e6e6')
        ax.plot(x,y, color = 'blue')
        canvas = FigureCanvas(fig)
        scene = QtWidgets.QGraphicsScene(self.graphicsView_4)
        scene.addWidget(canvas)
        self.graphicsView_4.setScene(scene)

    def get_year_balance(self):
        y = []
        x = ['янв','фев','мар','апр','май','июн','июл','авг','сен','окт','ноя','дек']
        P_max = np.max(self.table['Нагрузка'])
        for i in range(0,12):
            data = self.table.loc[self.table['Месяц']==(i+1)]
            F1 = 0
            F2 = 0
            for i in range(0,data.shape[0]):
                self.E_ey[i] = (data['Энергия_излуч.'].iloc[i]+data['Энергия_излуч.'].iloc[i-1])/2*self.S/1000000
                self.P[i] = (data['Нагрузка'].iloc[i]+data['Нагрузка'].iloc[i-1])/2
                dE = self.E_ey[i]*self.N_min-self.P[i]*self.Nu_i
                if dE>0:
                    F1+= dE*(self.a1*(dE-self.E_ey_nom) + self.b1)
                if dE<=0:
                    F2+= (-dE-self.E_vp)/(self.a2*(P_max-dE) + self.b2)+self.E_vp
            y.append(F1-F2)
        fig = Figure(figsize=(10,5))
        ax = fig.add_subplot()
        fig.patch.set_facecolor('#b6c3be')
        ax.set_title('Годовой баланс энергии, запасенной в водороде', fontsize = 12)
        ax.set_xlabel('Месяц', fontsize=10)
        ax.set_ylabel('dE, кВт', fontsize=10)
        ax.patch.set_facecolor('#e6e6e6')
        df = pd.DataFrame(data = y, columns = ['dE'])
        df['Месяц'] = x
        t = ""
        for i in range(0,12):
            t += df.iloc[i]['Месяц']+' '+str(np.round_(df.iloc[i]['dE'],2))+'\n'
        #ax.text(9, 1, t, ha='left', fontsize = 8, backgroundcolor = '#C0C0C0')
        ax.grid()
        ax.tick_params(axis = 'both', which = 'major', labelsize = 8)
        ax.plot(x,np.zeros(12),color = 'black')
        ax.plot(x,y, color = '#2a169c', label = t)
        ax.legend(loc = 'upper right')
        ax.scatter(x,y)
        canvas = FigureCanvas(fig)
        scene = QtWidgets.QGraphicsScene(self.graphicsView)
        scene.addWidget(canvas)
        self.graphicsView.setScene(scene)

    def Count_N_min(self):
        self.S = int(self.lineEdit.text())
        self.N_max = int(self.lineEdit_3.text())
        self.Nu_i = float(self.lineEdit_2.text())
        self.a1 = float(self.lineEdit_4.text())   #Nu_ey_x 0.0010332
        self.b1 = float(self.lineEdit_6.text())#1.422768  Nu_ey_c 0.73
        self.a2 = float(self.lineEdit_5.text())  #Nu_te_x 0.00087737
        self.b2 = float(self.lineEdit_7.text()) #Nu_te_c 0.5691056

        Max_Row = self.table.shape[0]
        #золотое сечение
        phi = (1+np.sqrt(5))/2
        right = self.N_max
        left = 0
        x2 = right - (right - left) / (phi+1)
        x1 = left + (right - left) / (phi+1)

        D1 = self.CalcD_n(self.table, x1)
        D2 = self.CalcD_n(self.table, x2)
        while np.abs(x1-x2)>0.5:
            if (D1<D2):
                right = x2
                x2 = x1
                D2 = D1
                x1 = left + (right - left) / (phi+1)
                D1 = self.CalcD_n(self.table, x1)
            elif (D1>=D2):
                left = x1
                x1 = x2
                D1 = D2
                x2 = right - (right - left) / (phi+1)
                D2 = self.CalcD_n(self.table, x2) 
        self.N_min=(left+right)/2
        self.N_min = np.ceil(self.N_min)
        
        self.label_31.setText(str(np.round(self.P_max,2)))
        self.label_32.setText(str(self.N_min))
        #добавление столбцов +dE -dE
        de_plus = []
        de_minus = []
        for i in range(0,Max_Row):
                self.E_ey[i] = self.table['Энергия_излуч.'].iloc[i]*self.S/1000000
                self.P[i] = self.table['Нагрузка'].iloc[i]
                dE = (self.E_ey[i]*self.N_min-self.P[i])
                if dE>0:
                    de_plus.append(np.round_((dE*self.Nu_i*(self.a1*(self.E_ey_nom-self.E_ey[i]*self.N_min)+self.b1)),4))
                    de_minus.append(0)
                elif dE<=0:
                    de_minus.append(np.round_((-dE/(self.Nu_i*(self.a2*(np.round_(self.P_max)-self.P[i])+self.b2))),4))
                    de_plus.append(0)
        self.table['+dE'] =de_plus
        self.table['-dE'] =de_minus
        #отрисовка
        for row in range(0,Max_Row):
            for column in range(6,8):
                item = QtWidgets.QTableWidgetItem()
                self.tableWidget.setItem(row, column, item)
                item.setText(str(self.table.iloc[row,column]))
        #отрисовка графика годового баланса
        self.get_year_balance()  
        self.get_daily_power(self.most_favorable_day()[1], self.most_favorable_day()[2])
        self.get_daily_workload(self.most_favorable_day()[1], self.most_favorable_day()[2])
        self.get_year_dE()
        #вывод переменных
        y1 = []
        y2 = []
        for i in range(0,Max_Row):
                self.E_ey[i] = self.table['Энергия_излуч.'].iloc[i]*self.S/1000000
                self.P[i] = self.table['Нагрузка'].iloc[i]
                y1.append(self.E_ey[i]*self.N_min-self.P[i])
                y2.append(self.E_ey[i]*self.N_min-self.P[i]*self.Nu_i)
        
        self.N_max_ey = np.max(y1)
        self.N1_te = np.abs(np.min(y2))
        self.label_11.setText(str(np.round_(self.N_max_ey,3)))
        self.label_35.setText(str(np.round_(self.N1_te,3)))

        F1 = 0
        F2 = 0
        y = []
        for i in range(0,12):
            data = self.table.loc[self.table['Месяц']==(i+1)]
            F1 = 0
            F2 = 0
            for i in range(0,data.shape[0]):
                self.E_ey[i] = (data['Энергия_излуч.'].iloc[i]+data['Энергия_излуч.'].iloc[i-1])/2*self.S/1000000
                self.P[i] = (data['Нагрузка'].iloc[i]+data['Нагрузка'].iloc[i-1])/2
                dE = self.E_ey[i]*self.N_min-self.P[i]*self.Nu_i
                if dE>0:
                    F1+= dE*(self.a1*(dE-self.E_ey_nom) + self.b1)
                if dE<0:
                    F2+= (-dE-self.E_vp)/(self.a2*(self.P_max-dE) + self.b2)+self.E_vp
            y.append(F1-F2)
        self.E_max_h2 = np.max(y)
        self.label_37.setText(str(np.round_(self.E_max_h2,3)))

    def Calc_ey_efficiency(self):
        try:
            A_v = float(self.lineEdit_8.text())
            y = []
            for i in range(0, self.table.shape[0]):
                y.append(self.table['Энергия_излуч.'].iloc[i]*self.S/1000000*self.N_min-self.table['Нагрузка'].iloc[i])
            self.label_33.setText(str(np.round_(np.max(y)/A_v,3))) 
            return np.max(y)/A_v
        except:
            print("Error")

    def Calc_ey_num(self):
        try:
            A_v = float(self.lineEdit_8.text())
            eff = float(self.lineEdit_9.text())
            v_h2 = self.Calc_ey_efficiency()
            self.label_34.setText(str(np.ceil(v_h2/eff)))
        except:
            print("Error")

    def Calc_te_num(self):
        try:
            n_te_nom = float(self.lineEdit_10.text())
            y = []
            for i in range(0, self.table.shape[0]):
                P = self.table['Нагрузка'].iloc[i]
                y.append((self.table['Энергия_излуч.'].iloc[i]*self.S/1000000*self.N_min-P*self.Nu_i))
            self.label_36.setText(str(np.ceil((np.abs(np.min(y)))/n_te_nom)))
        except:
            print("Error")    

    def Calc_num_bal_and_max_mass(self):
        try:
            p_h2 = float(self.lineEdit_11.text())
            v_bal = float(self.lineEdit_12.text())
            A_v = float(self.lineEdit_8.text())
            self.label_38.setText(str(self.Calc_max_mass(A_v)))
            self.label_39.setText(str(self.Calc_num_balloon(p_h2, v_bal, A_v)))
        except:
            print("Error")

    def Calc_max_mass(self, A_v):
        y = []
        for i in range(0,12):
            data = self.table.loc[self.table['Месяц']==(i+1)]
            F1 = 0
            F2 = 0
            for i in range(0,data.shape[0]):
                self.E_ey[i] = (data['Энергия_излуч.'].iloc[i]+data['Энергия_излуч.'].iloc[i-1])/2*self.S/1000000
                self.P[i] = (data['Нагрузка'].iloc[i]+data['Нагрузка'].iloc[i-1])/2
                dE = self.E_ey[i]*self.N_min-self.P[i]*self.Nu_i
                if dE>0:
                    F1+= dE*(self.a1*(dE-self.E_ey_nom) + self.b1)
                if dE<0:
                    F2+= (-dE-self.E_vp)/(self.a2*(self.P_max-dE) + self.b2)+self.E_vp
            y.append(F1-F2)
        ro_h2 = 0.089
        return np.round_(np.max(y)*ro_h2/A_v,3)

    def Calc_num_balloon(self, p_h2, v_bal, A_v):
        y = []
        for i in range(0,12):
            data = self.table.loc[self.table['Месяц']==(i+1)]
            F1 = 0
            F2 = 0
            for i in range(0,data.shape[0]):
                self.E_ey[i] = (data['Энергия_излуч.'].iloc[i]+data['Энергия_излуч.'].iloc[i-1])/2*self.S/1000000
                self.P[i] = (data['Нагрузка'].iloc[i]+data['Нагрузка'].iloc[i-1])/2
                dE = self.E_ey[i]*self.N_min-self.P[i]*self.Nu_i
                if dE>0:
                    F1+= dE*(self.a1*(dE-self.E_ey_nom) + self.b1)
                if dE<0:
                    F2+= (-dE-self.E_vp)/(self.a2*(self.P_max-dE) + self.b2)+self.E_vp
            y.append(F1-F2)
        return np.round_(np.max(y)/(p_h2*v_bal*A_v))
    
    def Calc_stoves(self):
        E_vp2 = 0
        try:
            t_stoves = float(self.lineEdit_13.text())
            n_hotplate = float(self.lineEdit_14.text())
        except:
            t_stoves = 0
            n_hotplate = 0
        S = self.lineEdit_15.text()
        if S!='':
            S = float(S)
            E_vp2 = S*0.1*24*30*7/2/self.table.shape[0]

        E_vp1 = 1.5*t_stoves*n_hotplate/24
        self.E_vp = E_vp1+E_vp2
        if not self.flag:
            self.table['Эл_нагрузка_']= self.table['Нагрузка']
            self.flag = True
        
        self.table['Нагрузка'] = self.table['Эл_нагрузка_'] + self.E_vp
        self.renew_table()
        self.Count_N_min()
        self.Calc_ey_efficiency()
        self.Calc_ey_num()
        self.Calc_te_num()
        self.Calc_num_bal_and_max_mass()
        self.label_48.setText(str(np.round_(1.5*t_stoves*n_hotplate/24*8784*0.42),2))

    def Calc_boiler(self):
        E_vp1 = 0
        try:
            S = float(self.lineEdit_15.text())
        except:
            S = 0
        t_stoves = self.lineEdit_13.text()
        n_hotplate = self.lineEdit_14.text()
        if t_stoves != '' and n_hotplate != '':
           t_stoves = float(t_stoves)
           n_hotplate = float(n_hotplate)
           E_vp1 = 1.5*t_stoves*n_hotplate/24
        E_vp2 = S*0.1*24*30*7/2/self.table.shape[0]
        self.E_vp = E_vp1+E_vp2
        if not self.flag:
            self.table['Эл_нагрузка_']= self.table['Нагрузка']
            self.flag = True
        self.table['Нагрузка'] = self.table['Эл_нагрузка_'] + self.E_vp
        self.renew_table()
        self.Count_N_min()
        self.Calc_ey_efficiency()
        self.Calc_ey_num()
        self.Calc_te_num()
        self.Calc_num_bal_and_max_mass()
        self.label_46.setText(str(np.round_(S*0.1*24*30*7/2*1.67),2))

def main(argv):
    app = QtWidgets.QApplication(argv)
    window = App()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main(sys.argv)        
