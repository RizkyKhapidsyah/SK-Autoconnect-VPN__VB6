Attribute VB_Name = "MWinInetDialUp"
'The Win32 Internet functions include
'seven functions that handle modem connections.

'------------------------------------------------------------------------------------
'Function Name: InternetAutodial
'------------------------------------------------------------------------------------
'Description: Automatically causes the modem to dial the default Internet connection.
'Returns: Returns TRUE if successful, or FALSE otherwise.
'Arguments
'dwFlags: Double-word value that contains the flags
'         controlling this operation. Can be one of
'         the following values:
'         INTERNET_AUTODIAL_FORCE_ONLINE  Forces an online Internet connection.
'         INTERNET_AUTODIAL_FORCE_UNATTENDED  Forces an unattended Internet dial-up.
'dwReserved: Reserved. Must be set to zero.
'------------------------------------------------------------------------------------
Public Declare Function _
InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, _
                                    ByVal dwReserved As Long) As Long
'------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
'Function Name: InternetAutodialHangup
'------------------------------------------------------------------------------------
'Goal: Disconnects an automatic dial-up connection.
'Returns: Returns TRUE if successful, or FALSE otherwise.
'Arguments
'dwReserved: Reserved. Must be set to zero.
'------------------------------------------------------------------------------------
Public Declare Function _
InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
'------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
'Function Name: InternetDial
'------------------------------------------------------------------------------------
'Goal: Initiates a connection to the Internet using a modem connection.
'Returns:
'Arguments
'hwndParent: Handle to the parent window.
'lpszConnectoid: String value containing the name of the
'                dial-up connection to use.
'dwFlags: Double-word value that contains the flags to use.
'         Can be one of the following values:
'         INTERNET_AUTODIAL_FORCE_ONLINE  Forces an online connection.
'         INTERNET_AUTODIAL_FORCE_UNATTENDED  Forces an unattended Internet dial-up.
'         INTERNET_DIAL_UNATTENDED  Connects to the Internet through a modem,
'                                   without displaying a user interface.
'lpdwConnection: Address of a double-word value containing the number
'                associated to the connection.
'dwReserved:     Reserved. Must be set to zero.

Public Declare Function _
InternetDial Lib "wininet.dll" (ByVal hwndParent As Long, _
                                ByVal lpszConnectoid As String, _
                                ByVal dwFlags As Long, _
                                lpdwConnection As Long, _
                                ByVal dwReserved As Long) As Long
'------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
'Function Name: InternetGetConnectedState
'------------------------------------------------------------------------------------
'Goal: Retrieves the connected state of the local system.
'Returns: Returns TRUE if there is an Internet connection, FALSE otherwise.
'Arguments
'lpdwFlags: Address of a double-word variable where the connection
'           description should be returned. Can be a combination of
'           the following values:
'            INTERNET_CONNECTION_MODEM       Local system uses a modem
'                                            to connect to the Internet.
'            INTERNET_CONNECTION_LAN         Local system uses a local
'                                            area network to connect to
'                                            the Internet.
'            INTERNET_CONNECTION_PROXY       Local system uses a proxy
'                                            server to connect to the Internet.
'            INTERNET_CONNECTION_MODEM_BUSY  Local system's modem is busy
'                                            with a non-Internet connection.

'dwReserved: Reserved. Must be set to zero.
'------------------------------------------------------------------------------------
Public Declare Function _
InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, _
                                             ByVal dwReserved As Long) As Long
'------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
'Function Name: InternetGoOnline
'------------------------------------------------------------------------------------
'Goal: Prompts the user for permission to initiate connection to a URL.
'Returns: Returns TRUE if successful, or FALSE otherwise.
'Arguments
'lpszURL:    String value containing the URL of the Web site to connect to.
'hwndParent: Handle to the parent window.
'dwReserved: Reserved. Must be set to zero.
'------------------------------------------------------------------------------------
Public Declare Function _
InternetGoOnline Lib "wininet.dll" (ByVal lpszURL As String, _
                                     ByVal hwndParent As Long, _
                                     ByVal dwReserved As Long) As Long
'------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------
'Function Name: InternetHangUp
'------------------------------------------------------------------------------------
'Goal: Instructs the modem to disconnect from the Internet.
'Returns:
'Arguments
'dwConnection: Double-word value that contains the number
'              assigned to the connection to be disconnected.
'dwReserved:   Reserved. Must be set to zero.
'------------------------------------------------------------------------------------
Public Declare Function _
InternetHangUp Lib "wininet.dll" (ByVal dwConnection As Long, _
                                  ByVal dwReserved As Long) As Long
'------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
'Function Name: InternetSetDialState
'------------------------------------------------------------------------------------
'Goal: Sets the modem dialing state.
'Returns: Returns TRUE if successful, or FALSE otherwise.
'Arguments
'lpszConnectoid: String value that contains the name of
'                the dial-up connection.
'dwState:        Double-word value that indicates the state
'                to set the dial-up connection to. Currently,
'                the only defined value is INTERNET_DIALSTATE_DISCONNECTED.
'dwReserved:     Reserved. Must be set to zero.

Public Declare Function _
InternetSetDialState Lib "wininet.dll" (ByVal lpszConnectoid As String, _
                                        ByVal dwState As Long, _
                                        ByVal dwReserved As Long) As Long
'------------------------------------------------------------------------------------


'// Flags for InternetDial - must not conflict with InternetAutodial flags
'//                          as they are valid here also.
Public Const INTERNET_DIAL_UNATTENDED = &H8000&  '0x8000

Public Const INTERENT_GOONLINE_REFRESH = &H1    '0x00000001
Public Const INTERENT_GOONLINE_MASK = &H1       '0x00000001

'// Flags for InternetAutodial
Public Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Public Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2
Public Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = 4


'// Flags for InternetGetConnectedState
Public Const INTERNET_CONNECTION_MODEM = 1
Public Const INTERNET_CONNECTION_LAN = 2
Public Const INTERNET_CONNECTION_PROXY = 4
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8

'// Flags for custom dial handler
Public Const INTERNET_CUSTOMDIAL_CONNECT = 0
Public Const INTERNET_CUSTOMDIAL_UNATTENDED = 1
Public Const INTERNET_CUSTOMDIAL_DISCONNECT = 2

'// Custom dial handler supported functionality flags
Public Const INTERNET_CUSTOMDIAL_SAFE_FOR_UNATTENDED = 1
Public Const INTERNET_CUSTOMDIAL_WILL_SUPPLY_STATE = 2
Public Const INTERNET_CUSTOMDIAL_CAN_HANGUP = 4

'// States for InternetSetDialState
Public Const INTERNET_DIALSTATE_DISCONNECTED = 1




