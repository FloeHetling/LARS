select 
[Юрлицо] as Company, 
[Office License Model] as OfficeLicenseModel, 
[Версия Офис] as OfficeVersion, 
[Windows License Model] as WindowsLicenseModel,
[Номер лицензии Windows OLP] as WindowsOLPSerial,
[Версия Windows] as WindowsVersion,
[Сетевое имя] as WSName,
[Серийный номер] as WSSerial
from aida.dbo.larspc where [Сетевое имя] = 'ws0006';