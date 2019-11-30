Option Explicit

On Error Resume Next

callingAlm

Sub callingAlm() 
	Dim xlApp 
  	Dim xlBook	
  	Set xlApp = CreateObject("Excel.Application") 
	Set xlBook = xlApp.Workbooks.Open(WScript.Arguments.Item(0), 0, True) 
  	xlApp.Run "QCConnection",WScript.Arguments.Item(1),WScript.Arguments.Item(2),WScript.Arguments.Item(3),WScript.Arguments.Item(4),WScript.Arguments.Item(5)
  	xlApp.Quit 

  	Set xlBook = Nothing 
  	Set xlApp = Nothing 

End Sub 