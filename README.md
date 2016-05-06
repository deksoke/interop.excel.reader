# ExcelReader
Using Microsoft.Office.Interop.Excel 
Read Excel To DataTable (vb.net)

# ความต้องการ (Requirement)
+ ติดตั้ง MS Office 2010 หรือเวอร์ชั่นสูงกว่า (2010 or Above)
+ Import ExcelReaderPlugin.ExcelReader ใส่ในคลาสที่ต้องการเรียกใช้

# เริ่มต้นการใช้งาน
VB
```vb
Dim FilePath as string = "C:/201604.xlsx"
Using xlsReader As New ExcelReader(FilePath)
    dsWorkBook = xlsReader.GetDataSet(True) //Return DataSet
    dtWorkbook = xlsReader.GetDataTable(0) //Return DataTable
    dtSheetLists = xlsReader.GetSheetLists() //Return DataTable
    SheetCount = xlsReader.GetSheetCount() //Return Integer
End Using
```

# ฟังชั่น
Methods | Parameters | Return Type |Description
---------|------------| ---------|---------------
GetSheetLists | - | DataTable | ชื่อชีททั้งหมด
GetSheetCount | - | Integer | จำนวนชีท
GetDataSet | SetColumnNameWithColumnHeader as boolean = false, AutoChangeDataType as boolean = false  | DataSet | รีเทิร์นข้อมูลทั้งหมดในไฟล์ excel หลังจากที่อ่านและแปลงออกมาให้อยู่ในรูปของ DataSet
GetDataTable | tbIndex as integer, SetColumnNameWithColumnHeader as boolean = false, AutoChangeDataType as boolean = false | DataTable | รีเทิร์นข้อมูลเฉพาะชีทที่เลือก โดยแปลงให้อยู่ในรูปแบบของ DataTable แล้ว

หมายเหตพารามิเตอร์
+ SetColumnNameWithColumnHeader 
//true กำหนดชื่อคอลัมน์จากข้อมูลแถวแรก, false ตั้งเป็น column1 column2 column3
+ AutoChangeColumnDataType 
//true เปลี่ยนประเภทของคอลัมน์ตามข้อมูลจากแถวแรก, false จะกำหนดให้เป็น String ไว้ก่อน

#ประเภทไฟล์ที่รองรับ (support file extensions)
+ xls
+ xlsx
+ csv


# คัดลอกและปรับปรุงจาก (Code Reference)
