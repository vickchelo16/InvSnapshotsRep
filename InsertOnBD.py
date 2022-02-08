from time import strftime 
from datetime import datetime
from unittest import result
from MySQLdb import Connect 
import pandas as pd 
import pyodbc   

class Excel:
    def __init__(self,name,sheetname):
        self.name = name
        self.sheetname = sheetname

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):        
        return self
    
    @property
    def nameproperty(self):
        return self.name
    
    @property
    def sheetnameproperty(self):
        return self.sheetname

    def executeexcel(self):
        print('Reading excel: '+ str(datetime.now()))
        rd = pd.read_excel (self.name, sheet_name=self.sheetname)      
        print('Excel done:'+ str(datetime.now()))
        return rd   

class Errors:
    def __init__(self,error,tag):
        self.displayMessage(error,tag)

    def displayMessage(self,Message,Error):
        print(Message + "-"+ Error)

class Connection:
    def __init__(self):
        server = 'tcp:testserverisc.database.windows.net'
        port = '1433'
        Database = 'Snapshots'
        Uid = 'testserverisc'
        Password = '******'
        Timeout = '30'
        self._conn =  cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};Server='+server+','+port+';Database='+Database+';Uid='+Uid+';Pwd='+Password+';Encrypt=yes;TrustServerCertificate=no;Connection Timeout='+Timeout+';')  
        self._cursor = self._conn.cursor()

    def __enter__(self):#Enter, default method
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    @property
    def connection(self):
        return self._conn

    @property
    def cursor(self):
        return self._cursor

    def commit(self):
        self.connection.commit()

    def close(self, commit=True):
        if commit:
            self.commit()
        self.connection.close()

    def execute(self, sql, params=None):
        self.cursor.execute(sql, params or ())

    def fetchall(self):
        return self.cursor.fetchall()

    def fetchone(self):
        return self.cursor.fetchone()

    def query(self, sql, params=None):
        self.cursor.execute(sql, params or ())
        return self.fetchall()

class Business:   
    def vInsertOnBD(self,strFile,sheetname):    
        with Excel(strFile,sheetname) as excel:
            df = excel.executeexcel() 
            print (df) 
        for index,row in df.iterrows():
            material = row['Material_ID']
            site = row['Site_ID']
            source = row['Source System']
            bu = row['BU']
            region = row['Region']
            invqty = row['Sum of InvQuantity In EA']
            invvalue = row['Sum of Inv Value Overall']
            sRes = self.vInsert(material,site,source,bu,region,invqty,invvalue) 
        print('Execution ended')
    def vSelect(self):
        try: 
            with Connection() as db: 
                results = db.query('SELECT * FROM SnapshotTable')
                print(results) 
        except Exception as e:
            Errors(str(e),'ERROR') 

    def vInsert(self,material,site,sourcesystem,bu,region,invqty,invvalue):  
            try:  
                with Connection() as db:
                    now = datetime.now()
                    sDate = now.strftime("%Y-%m-%d %H:%M:%S")  
                    db.execute('Insert into [dbo].[SnapshotTable](Material_ID, Site_ID,Source_System,BU,Region,SnapshotDate,InvQty,InvValue) values(?,?,?,?,?,?,?,?)',(material,site,sourcesystem,bu,region,sDate,invqty,invvalue))
                    return 'OK'           
            except Exception as e:
                Errors(str(e),'ERROR') 


Business().vSelect()
Business().vInsertOnBD(r"C:\Users\10305087\Downloads\20220124IHD_Weekly_Snapshot.xlsx","Sheet1")



    
