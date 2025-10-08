import pandas

# RULES:
# 1. JANGAN GANTI NAMA CLASS ATAU FUNGSI YANG ADA
# 2. JANGAN DELETE FUNGSI YANG ADA
# 3. JANGAN DELETE ATAU MENAMBAH PARAMETER PADA CONSTRUCTOR ATAU FUNGSI
# 4. GANTI NAMA PARAMETER DI PERBOLEHKAN
# 5. LARANGAN DI ATAS BOLEH DILANGGAR JIKA ANDA TAU APA YANG ANDA LAKUKAN (WAJIB BISA JELASKAN)
# GOODLUCK :)

class excelManager:
    def __init__(self,filePath:str,sheetName:str="Sheet1"):
        self.__filePath = filePath
        self.__sheetName = sheetName
        self.__data = pandas.read_excel(filePath,sheet_name=sheetName)
            
    def insertData(self,newData:dict,saveChange:bool=False):
        # Menambahkan baris baru ke DataFrame
        newRow = pandas.DataFrame([newData])
        self.__data = pandas.concat([self.__data, newRow], ignore_index=True)
        
        # Simpan perubahan jika diminta
        if (saveChange): self.saveChange()

    def deleteData(self, targetedNim:str,saveChange:bool=False):
        # Cari baris berdasarkan NIM
        indexToDelete = None
        for i in self.__data.index:
            if str(self.__data.at[i, "NIM"]) == str(targetedNim):
                indexToDelete = i
                break
        
        # Hapus baris jika ditemukan
        if indexToDelete is not None:
            self.__data.drop(indexToDelete, inplace=True)
            self.__data.reset_index(drop=True, inplace=True)
        
        # Simpan perubahan jika diminta
        if (saveChange): self.saveChange()

    def editData(self, targetedNim:str, newData:dict,saveChange:bool=False) -> dict:
        # Cari baris berdasarkan NIM
        indexToEdit = None
        for i in self.__data.index:
            if str(self.__data.at[i, "NIM"]) == str(targetedNim):
                indexToEdit = i
                break
        
        # Edit data jika ditemukan
        if indexToEdit is not None:
            for key in newData:
                if key in self.__data.columns:
                    self.__data.at[indexToEdit, key] = newData[key]
            
            # Simpan perubahan jika diminta
            if (saveChange): self.saveChange()
            
            # Kembalikan data yang sudah diedit
            return {col: str(self.__data.at[indexToEdit, col]) for col in self.__data.columns}
        
        return None
                    
    def getData(self, colName:str, data:str) -> dict:
        collumn = self.__data.columns
        
        collumnIndex = [i for i in range(len(collumn)) if (collumn[i].lower().strip() == colName.lower().strip())] 
        
        if (len(collumnIndex) != 1): return None
        
        colName = collumn[collumnIndex[0]]
        
        resultDict = dict()
        
        for i in self.__data.index:
            cellData = str(self.__data.at[i,colName])
            if (cellData == data):
                for col in collumn:
                    resultDict.update({str(col):str(self.__data.at[i,col])})
                resultDict.update({"Row":i})
                return resultDict
        
        return None
    
    def saveChange(self):
        self.__data.to_excel(self.__filePath, sheet_name=self.__sheetName , index=False)
    
    def getDataFrame(self):
        return self.__data
