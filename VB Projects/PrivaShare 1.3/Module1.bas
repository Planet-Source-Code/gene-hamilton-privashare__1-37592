Attribute VB_Name = "Module1"
Option Explicit

'Used by the timer routine to know who to send myInfo to.
Global intChannel As Integer

'Used to select text for color change.
Global intSelText As Integer

'Put app.path in a string
Global appPath As String

'******** Connection variables *********************

'Information for each contact.
Public Type connection
    Name As String
    IP As String
    Sharing As Boolean
    relay As Boolean
End Type

'Make the collection.
Public Connect() As connection

'Make temporary node to add to node tree view.
Global nodTemp As Node

'Number of connections in array.
Global intNum_Connections As Integer

'Number of connections in use.
Global intNum_ConnectionsNow As Integer

'Port address for default connectios.
Global intPort As Integer

'Remember who your getting file from, in case it stops sending.
Global memoryIndex As Integer
Global memoryChannel As Integer

'Interger array for multiple downloads.
'Each file has it's own channel. #2-202 to save and read from.
'#1 is used for saving logs and preferences.
Global intChannels(2 To 202) As Integer

'Logging turned on?
Global blnLog As Boolean

'Welcome string shown when someone connects to you.
Global strWelcome As String

'Make string for file transfers.
Global strFileString As String

'If you want privashare to be passive and not interupt
'gaming, press the passive button.
Global passive As Boolean

'Temp information strings
Global strName As String
Global strIP As String
Global strIndex As String
Global blnShare As Boolean

'Value is true if file being sent is a sound.
Global blnWav As Boolean

' Security  variables.
Global blnSecure As Boolean 'true if password on, false if off.
Global intStrikes As Integer 'Number of times allowed to try at your password.
Global intAccess() As Integer 'Dynamic array. 1 if accepted, negitive if strikes. Posible access levels later.
Global strPassword As String 'Your servers password.
