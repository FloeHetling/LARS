select 
[������] as Company, 
[Office License Model] as OfficeLicenseModel, 
[������ ����] as OfficeVersion, 
[Windows License Model] as WindowsLicenseModel,
[����� �������� Windows OLP] as WindowsOLPSerial,
[������ Windows] as WindowsVersion,
[������� ���] as WSName,
[�������� �����] as WSSerial
from aida.dbo.larspc where [������� ���] = 'ws0006';