Attribute VB_Name = "CheckifDunconnect"

Declare Function RasEnumConnections Lib "RasApi32.DLL" Alias "RasEnumConnectionsA" (lprasconn As Any, lpcb As Long, lpcConnections As Long) As Long

Function RASCount() As Integer
  Dim lprasconn(0 To 1) As Long ' dummy buffer area
  Dim rc As Long ' return code
  Dim lpcb As Long ' buffer size
  Dim lpcConnections As Long ' connection count

  lprasconn(0) = 32 ' each returned item is at least 32 bytes long
  lpcb = 0 ' set total number of usable bytes in the buffer to zero
  
  rc = RasEnumConnections(lprasconn(0), lpcb, lpcConnections)
  RASCount = lpcConnections ' return connection count

End Function


