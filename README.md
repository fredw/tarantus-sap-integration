SAP Dll Location
================================================
C:\Program Files (x86)\SAP\SAP Business One Server\B1_SHR\DIAPI.x64\Program Files 64\SAP\SAP Business One DI API\DI API 90


How to fix error to add SAPbobsCOM.dll as Reference
================================================
When you try to add **SAPbobsCOM.dll** on references and get the error:

	A reference to the "....dll" could not be added.  
	Please make sure that the file is accessible and that it is a valid assembly or COM component.

Solution:

1. Run `cmd` as administrator/  
2. Go to folder `cd "C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin"` (if this folder doesn't exists, try to find him: `dir tlbimp.exe /s`)  
3. Add dll: `TlbImp.exe C:\YourDllFolder\SAPbobsCOM.dll`  
4. A new dll has been created in the same folder of TlbImp.exe. You can use that as reference in you project.

Source: [http://stackoverflow.com/questions/3456758/a-reference-to-the-dll-could-not-be-added%60](http://stackoverflow.com/questions/3456758/a-reference-to-the-dll-could-not-be-added%60)