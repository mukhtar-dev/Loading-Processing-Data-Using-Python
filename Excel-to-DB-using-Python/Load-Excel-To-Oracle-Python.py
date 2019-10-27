'''
Created on Sep 9, 2019

@author: m.isameldin
'''
import pandas as pd
import cx_Oracle
import time
from tkinter import *
from tkinter.ttk import *
import collections

 
    

class Window(Frame):
    arr_val ={}
    tabl_clmn = []
    sht_name ='DB'
    tg_tabl = 'CELL_TEMP'
    #tg_tabl = 'CELL_LOOKUP_TMP'

    def __init__(self,master=None):
        Frame.__init__(self, master)                 
        self.master = master
        #self.init_window()
                # changing the title of our master widget      
        self.master.title("GUI")

        # allowing the widget to take the full space of the root window
        self.pack(fill=BOTH, expand=1)

        # creating a button instance
        quitButton = Button(self, text="Load The Data",command =Window.helloCallBack)

        # placing the button on my window
        quitButton.place(x=300, y=0)
        
        self.db = cx_Oracle.connect('ocdm_test/oraoctest_02@ocdm11g')
        self.cl = self.db.cursor()
        self.cl.execute('select /*+ parallel(20) */  column_name from user_tab_columns where table_name = :tb ',tb=self.tg_tabl)
        df1 = pd.read_excel(r'E:\work\CELLs_CHECK\Lookup_Sep_2019.xlsx', sheet_name=self.sht_name)  
        #print(df1.columns)
        df_c = []
        for df_clmn in df1.columns:
            df_c.append(df_clmn)
        print(df_c)

        self.combos=[]
        self.labels=[] #creates an empty list for your labels
        for x in self.cl: #iterates over your nums 
            self.label = Label(root,text=x) #set your text
            self.label.pack()
            self.tabl_clmn.append(x[0])
            #label.place(x=xx,y=yy)
            self.labels.append(self.label) #appends the label to the list for further use
            self.combo = Combobox(root)
            self.combo['values']= df_c
            #combo.current(1) #set the selected item
            self.combo.pack()
            #self.combo.bind("<<ComboboxSelected>>", lambda event,x=x : self.justamethod(self.combo)) 
            self.combos.append(self.combo)
        odr = 0
        for cm in self.combos:
            cm.bind("<<ComboboxSelected>>", lambda event,cm=cm,odr=odr : self.justamethod(cm,odr))
            odr = odr + 1
        self.cl.close()
        self.db.close()
        print(self.tabl_clmn)
           

        
    def justamethod (self,eventObject,i):
        
        print("method is called")
        #print (eventObject.get(),eventObject)
        self.arr_val[i]=eventObject.get()
    
             
    def helloCallBack():   
        print('the button has been selected ')
        sel_clm =Window.tabl_clmn
        val =[]
        sel_clm2=[]
        val_dict = Window.arr_val
        od = collections.OrderedDict(sorted(val_dict.items()))
        for k, v in od.items():
            val.append(v)
            sel_clm2.append(sel_clm[k])
        print(val)
        #val = ['dept','name']
        dfrm=read_ex(r'E:\work\CELLs_CHECK\Lookup_Sep_2019.xlsx',Window.sht_name)
        print(dfrm.columns)
        static_val ={}#{1:'MyName'} This should be replaced by the static values to be inserted if any 
        ora_table(dfrm,sel_clm2,val,static_val,Window.tg_tabl)
        print(dfrm.columns)  
        print("--- The Application Run time was: %s seconds ---" % (time.time() - start_time))
        print(len(sel_clm))

    #Creation of init_window
 #   def init_window(self):



def read_ex(f_name,sh):   # function to read excel file , and list the value of sheet no 1
        xls = pd.ExcelFile(f_name)                                   # read the excel files
        ''' sheets=xls.sheet_names
        sh1 = sheets[0]
        print('Sheets In this excel file are: ',sh1)          # get all sheets name within the files 
        print('___________________________________')'''
        df1 = pd.read_excel(xls, sheet_name=sh)                         # load the data from one sheet 
    
        #print(df1)
                                      # print the sheet data         
        
        return(df1)    

def ora_table(dfrm1,clmns,val1,st_val,tabl):

        db = cx_Oracle.connect('ocdm_test/oraoctest_02@ocdm11g')
        cl = db.cursor()
        
  
        #print(clmns)
        
        stms = 'insert into ' + tabl +'('
        stms = stms + ','.join(clmns)
        stms = stms+') values(:'
        stms = stms + ',:'.join(clmns)
        stms = stms + ')'
        print(stms) 
        
        #print('insert /*+ append */  into {tb} ({cl1}, {cl2}, {cl3}, {cl4}) values (:v1,:v2,:v3,:v4)'.format(tb='PY_EMP',cl1=clmns[0],cl2=clmns[1],cl3=clmns[2],cl4=clmns[3]))
        dfrm1 = dfrm1.loc[:,val1]
        for idx,row in dfrm1.iterrows(): 
            lst = []
            
            
            for r in row:
                lst.append(r)
            for key in st_val:
                lst.insert(key, st_val[key])
            print(lst)
            cl.execute(stms,lst)
            print('1 row Loaded...')
            db.commit()
            
        '''    
        for idx,row in dfrm1.iterrows(): 
            cl.execute('insert /*+ append */  into {tb} ({cl1}, {cl2}, {cl3}, {cl4}) values (:v1,:v2,:v3,:v4)'.format(tb='PY_EMP',cl1=clmns[0],cl2=clmns[1],cl3=clmns[2],cl4=clmns[3])
                       ,v1=row['id'],v2=row['name'],v3=row['salary'],v4=row['dept'])
            print('1 row Loaded...')
            db.commit()
            '''

        
        cl.close()
        db.close()
        

        
if __name__ == '__main__':
    start_time = time.time()

    root = Tk()
    root.geometry("400x700")
    app = Window(root)
    
    root.mainloop()
      
  
 
    

