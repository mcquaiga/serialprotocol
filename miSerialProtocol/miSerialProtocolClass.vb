Imports System.Windows.Forms

'***************
'Author: Adam McQuaig
'Date: Oct 5, 2007
'Description: This is a bass class for all the instruments in the mercury product line,
'               each seperate instrument must inherit from this class and then functionality can be 
'               extended to each specific instrument as needed
'   I hope this will also reduce the amount of code and guess work
'   
'   There was no need for auto detection so I can call the specific instrument based on what the user chooses,
'   a inherited class can be created to auto detect the instrument

'To do:
' 1. Better error handling - As of now if no instrument is hooked up and we try to connect to it, a proper response is received but component are not
' being released properly
' 2. Live Read Functions
' 3. CRC checking on RD and RG functions

#Region "Enums"

Public Enum BaudRateEnum
    'All the instrument baud rates
    b300 = 300
    b600 = 600
    b1200 = 1200
    b2400 = 2400
    b4800 = 4800
    b9600 = 9600
    b19200 = 19200
    b38400 = 38400
End Enum
Public Enum InstrumentTypeCode
    Mini = 2
    ECAT = 0
    MiniAT = 3
    MiniMax = 4
    AccuTest = 21
    TurboCor = 20
    TurboMon = 6
    PulseAcc = 8
    TCI = 12
    ER = 1
    ERX = 10
    PTBoard = 5
    AutoDetect = 99
End Enum
Public Enum CommStateEnum
    UnlinkedIdle
    IncomingAwaitingSITE
    IncomingConnected
    WakingItUp
    SigningOn
    LinkedIdle
    ReadingItem
    ReadingGroup
    ReadingFirstLog
    ReadingMoreLogs
    WritingItem
    SigningOff
    InitializingModem
    Listening
    SendingClearAlarmsMessage
    Dialing
    HangingUp
End Enum
Public Enum MessageStateEnum
    OK_Idle
    WaitingForReceiveEvent
    ReceivedAPacket
    ProtocolError
    FramingError
    ReceivedERROR
    OtherError
End Enum
Public Enum CommCharEnum
    SOH = 1
    STX = 2
    ETX = 3
    EOT = 4
    ENQ = 5
    ACK = 6
    CR = 13
    NAK = 21
    RS = 30
End Enum
Public Enum InstrumentErrorsEnum

    NoData = -1
    NoError = 0
    MessageFormatError = 1
    SignOnError = 20 'Wrong Instrument Type
    TimeoutError = 21 'Timeout Error
    FramingError = 22 'Framing Error
    CheckByteError = 23 'Check byte error
    InvalidAccessCodeError = 27 'Invalid Access Code Error
    InvalidCommandError28 'Invalid Command Error
    InvalidItemNumberError = 29 'Invalid Item Number
    InvalidEnquiryError = 30 'Invalid Enquiry
    TooManyRetransmissionsError = 31 'Too many retransmissions
    ReadOnlyError = 32 'Attempted access is disabled. item is read-only
    NoAuditTrailError = 33 'No audit trail
    EventLoggerFullError = 34 'Event logger full
    InvalidAGA8Condition = 35
    InvalidDataError = 38
    InvalidChangeAttempted = 39
    InsufficientPower = 40
    REIError = 41

End Enum
Public Enum AlarmValuesEnum
    Alarm = 11111111
    NoAlarm = 0
End Enum

#End Region

Public MustInherit Class miSerialProtocolClass
    Implements IDisposable

    'Using the new SerialPort class that comes with VB 2005, it seems to be alot more robust then the
    'previous MSComm Control
    Protected WithEvents comm As System.IO.Ports.SerialPort

    'Communication Port Settings
    'These must be set before calling connect or trying to open the serial port
    Protected i_CommPort As String
    Protected i_BaudRate As BaudRateEnum
    Protected i_Timeout As Integer
    Protected i_AccessCode As Integer = 33333

    'Comm State holds the higher level functions that are currently occuring
    Private i_CommState As CommStateEnum
    'MessageState holds the lower level functions that are currently occuring
    'i.e. Data Packets that are being recieved and if a message is in the midst of being sent
    Private i_MessageState As MessageStateEnum
    'Instrument that is currently being connected to
    'This is currently being set by the child class
    Private i_InstrumentType As InstrumentTypeCode
    Private i_InstrumentError As InstrumentErrorsEnum

    Private TimedOut As Boolean

    Private i_CommBuffer As String
    Private CommBuffer_position As Integer = 0 'The current cell in the array to return 
    Private comm_buffer As New Collection
    Public disposed As Boolean = False

    'Raised events when different states are changed
    'This will allow for logging capabilities at a later date
    Public Shared Event CommStateChanged(ByVal State As CommStateEnum)
    Public Shared Event CurrentItemProgress(ByVal NumItems As Integer)
    Public Shared Event ItemCount(ByVal Count As Integer)
    Private i_ItemsToDownload As Integer

    Protected Shared Event CommBufferChanged()
    Protected Shared Event MessageStateChanged()
    Protected Shared Event InstrumentErrorChanged()



#Region "Constructor and Destructor"

    Sub New(ByVal PortNumber As Integer, ByVal BaudRate As BaudRateEnum, Optional ByVal Timeout As Integer = 50)
        Try
            Me.Port = "COM" & PortNumber
            Me.BaudRate = BaudRate
            Me.Timeout = Timeout
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub

    Public Overloads Sub Finalize()
        Me.Dispose(True)
    End Sub

    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        ' Take yourself off of the finalization queue
        ' to prevent finalization code for this object
        ' from executing a second time.
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        ' Check to see if Dispose has already been called.
        If (disposing) Then
            'If diposing we must safely disconnect from the instrument if theres a connection
            If i_CommState = CommStateEnum.LinkedIdle Then
                Try
                    Me.Disconnect()
                Catch ex As Exception
                    Console.Write(ex.Message)
                End Try
            End If
            ' Dispose managed resources.
            If Not comm Is Nothing Then
                Console.Write("Disposing Serial Port Resources." & vbNewLine)
                comm.Close()
                comm = Nothing
            End If
        End If
    End Sub
#End Region

#Region "Properties"

    Public Property Port() As String
        Get
            Return i_CommPort
        End Get
        Set(ByVal value As String)
            i_CommPort = value
        End Set
    End Property

    Public Property AccessCode() As Integer
        Get
            Return i_AccessCode
        End Get
        Set(ByVal value As Integer)
            i_AccessCode = value
        End Set
    End Property

    Public Property BaudRate() As BaudRateEnum
        Get
            Return i_BaudRate
        End Get
        Set(ByVal value As BaudRateEnum)
            i_BaudRate = value
        End Set
    End Property

    Public Property Timeout() As Integer
        Get
            Return i_Timeout
        End Get
        Set(ByVal value As Integer)
            i_Timeout = value
        End Set
    End Property

    Protected Property CommState() As CommStateEnum
        Get
            Return i_CommState
        End Get
        Set(ByVal value As CommStateEnum)
            i_CommState = value
            RaiseEvent CommStateChanged(value)
            Console.Write(i_CommState.ToString & vbNewLine)
        End Set
    End Property

    Protected Property MessageState() As MessageStateEnum
        Get
            Return i_MessageState
        End Get
        Set(ByVal value As MessageStateEnum)
            i_MessageState = value
            RaiseEvent MessageStateChanged()
            Console.Write(i_MessageState.ToString & vbNewLine)
        End Set
    End Property

    Protected Property Instrument() As InstrumentTypeCode
        Get
            Return i_InstrumentType
        End Get
        Set(ByVal value As InstrumentTypeCode)
            i_InstrumentType = value
        End Set
    End Property

    'This property is what should be accessed when checking instrument errors
    'The WD function will set this value after it has sent a command to the instrument
    Public Property InstrumentError() As InstrumentErrorsEnum
        Get
            Return i_InstrumentError
        End Get
        Set(ByVal value As InstrumentErrorsEnum)
            'I'm not sure if it would be better or not to have this procedure throw an exception should the instrument return an error,
            'I'll leave it for now but it might make more sense

            i_InstrumentError = value
            RaiseEvent InstrumentErrorChanged()
            Console.Write(InstrumentError.ToString & vbNewLine)
        End Set
    End Property

    'The following properties will hold on to the incoming Comm data until it is used
    'I would like the two properties to control the rediming of the array and the deletion of the used items

    'As information is recieved into the serial port is it saved to the last available spot in the array, we must redim the array as needed
    'After a string has been requested it is deleted from the array and everything is bumped up a spot in the stack
    'Therefore a single request to the comm buffer will result in the deletion of the previous data so one must be careful
    'The best thing to do is to assign the value to a local variable and use it as seen fit

    'As of right now because the Comm Event handler is not being used, an item will be added to the comm buffer and used immediately

    'This will also allow us to log all the incoming serial data, it has not been implemented yet
    Protected Property CommBuffer() As String
        Get
            Dim data As String
            Dim x As Integer = 0

            data = comm_buffer.Item(comm_buffer.Count)
            If comm_buffer.Count >= 1 Then
                Console.Write("Retrieved " & data & " from Comm Buffer.")
                Return data
            Else
                Console.Write("Buffer is Empty." & vbNewLine)
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            If value <> Nothing And value <> "" Then
                comm_buffer.Add(value)
                RaiseEvent CommBufferChanged()
                Console.Write("Added " & value & " to Buffer." & vbNewLine)
            End If
        End Set
    End Property

    Public Property ItemToDownload() As Integer
        Get
            Return i_ItemsToDownload
        End Get
        Set(ByVal value As Integer)
            i_ItemsToDownload = value
            RaiseEvent ItemCount(value)
        End Set
    End Property

#End Region

#Region "Methods"

#Region "Public"
    'I hope this method will connect to any instrument
    'but if not we can override the method from within the inherited class
    '
    '
    Public Overridable Sub Connect()
        Dim instr_Code As String = Instrument
        Dim return_code As InstrumentErrorsEnum
        Try
            OpenCommPort()
            ClearCommBuffer()
            If CommState = CommStateEnum.UnlinkedIdle Then
                MessageState = MessageStateEnum.OK_Idle
                CommState = CommStateEnum.WakingItUp
                'After send EOT and ENQ we can expect a one character response from the instrument
                'Wake up the instrument
                comm.Write(Chr(CommCharEnum.EOT))
                System.Threading.Thread.Sleep(150)
                comm.Write(Chr(CommCharEnum.ENQ))
                RcvDataFromComm()
                return_code = GetReturnCode()
                'Attempt to connect if there was no error or if we get an Invalid Enquiry(30) back from the instrument, which means the instrument is stuck
                ' in a connection already
                If return_code = InstrumentErrorsEnum.NoError Or return_code = InstrumentErrorsEnum.InvalidEnquiryError Then
                    If Instrument < 10 Then instr_Code = "0" & Instrument
                    CommState = CommStateEnum.SigningOn
                    System.Threading.Thread.Sleep(100)
                    SendDataToComm("SN," & i_AccessCode & Chr(CommCharEnum.STX) & "vq" & instr_Code & Chr(CommCharEnum.ETX))
                    return_code = GetReturnCode()
                    If return_code = InstrumentErrorsEnum.NoError Then
                        CommState = CommStateEnum.LinkedIdle
                    Else
                        Throw New InstrumentCommunicationException(return_code)
                    End If
                ElseIf return_code = InstrumentErrorsEnum.NoData Then
                    'No data recieved, is an instrument hooked up?
                    If MessageBox.Show("No response from instrument." & vbNewLine & "Try Again?", "Connection Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error) Then
                        CommState = CommStateEnum.UnlinkedIdle
                        Me.Connect()
                    Else
                        Throw New CommExceptions("No Response From Instrument.")
                    End If
                Else
                    Throw New CommExceptions("An error occured: " & InstrumentError)

                End If
            ElseIf CommState = CommStateEnum.LinkedIdle Then
                'Already Connected to instrument, don't need to do anything
                Console.Write("Already Connected.")
            Else
                Throw New InstrumentBusyException(CommState)
            End If
        Catch ex As NoDataRecievedException
            CommState = CommStateEnum.LinkedIdle
        Catch ex As CommInUseException
            Throw New CommInUseException(comm.PortName)
        Catch ex As CommExceptions
            Me.Dispose()
            Throw New CommExceptions(ex.Message)
        Catch ex As TimeoutException
            If MessageBox.Show("Connection Timed Out." & vbNewLine & "Try Again?", "Failed Connection.", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error) = DialogResult.Retry Then
                CommState = CommStateEnum.UnlinkedIdle
                Me.Connect()
            Else
                Throw New CommExceptions("Connection cancelled.")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub

    'Signing off is common across all instrument types, but inherited classes can override the function if
    'needed, otherwise they can shadow it and do any clean up that is specific to that instrument
    'Calling the disconnect function will also tirgger the dispose function, freeing up the comm port
    Public Overridable Sub Disconnect()
        Dim return_code As InstrumentErrorsEnum

        If CommState = CommStateEnum.LinkedIdle Then
            CommState = CommStateEnum.SigningOff
            SendDataToComm("SF" & Chr(CommCharEnum.ETX))
            return_code = GetReturnCode()
            If return_code = InstrumentErrorsEnum.NoError Then
                CommState = CommStateEnum.UnlinkedIdle
            End If

            'Give the instrument half a second 
            System.Threading.Thread.Sleep(500)
            Me.Dispose()
        End If
    End Sub

    'Used to write values to the instrument
    'We can expect a simple message back from the instrument, with either an error or a success code

    'This function will take in an ItemNumber and a Value and write it to the instrument
    'All we need to do here is check the status code that the instrument sent back
    'Eventually I'm going to throw the error code off to a custom exception handler, this class will then
    'handle any errors as seen fit
    Public Function WD(ByVal ItemNumber As Integer, ByVal ItemValue As String) As InstrumentErrorsEnum

        Dim cmd As String

        Try
            'Check for connection to instrument
            'Throw an exception if we're not and let who's ever calling this function deal with it
            If i_CommState <> CommStateEnum.UnlinkedIdle Or i_CommState <> CommStateEnum.HangingUp Then
                CommState = CommStateEnum.WritingItem
                cmd = "WD," & i_AccessCode & Chr(CommCharEnum.STX) & ItemNumber.ToString.PadLeft(3, "0") & "," & ItemValue.PadLeft(8, "0") & Chr(CommCharEnum.ETX)
                SendDataToComm(cmd)


                CommState = CommStateEnum.LinkedIdle
                Return GetReturnCode()
            Else
                Throw New NoPhysicalConnectionException
            End If
        Catch ex As NoPhysicalConnectionException
            Throw New NoPhysicalConnectionException
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    'Single read operation to the instrument
    'It will return either a string with the value returned or if the item number is invalid we'll receive a InstrumentErrorsEnum object
    '
    'As of right now this function does no error handling, it simply returns the error that it saw from the instrument
    Public Function RD(ByVal ItemNumber As Integer) As Object
        Dim cmd As String
        Dim i_Object As Object
        Dim i_Error As InstrumentErrorsEnum
        Dim i_String As String = ""
        Dim i_Double As Double

        Try
            comm.DiscardInBuffer()
            CommState = CommStateEnum.ReadingItem
            cmd = "RD" & Chr(CommCharEnum.STX) & ItemNumber.ToString.PadLeft(3, "0") & Chr(CommCharEnum.ETX)
            SendDataToComm(cmd)
            i_Object = ParseRD()
            CommState = CommStateEnum.LinkedIdle
            'If an enum is returned then we have an error
            If i_Object.GetType.IsEnum = True Then
                i_Error = i_Object
                Throw New InstrumentCommunicationException(i_Error, ItemNumber)
            ElseIf i_Object.GetType.IsClass = True Then
                Try
                    i_Double = CDbl(i_Object)
                    Return i_Double
                Catch ex As Exception
                    i_String = i_Object
                    Return i_String
                End Try
            Else
                Return i_Object
            End If
        Catch ex As InvalidItemNumberException
            MessageBox.Show(ex.Message & vbNewLine & "Item Number: " & ex.ItemNumber)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function RG(ByVal Items As Collection, ByRef Downloaded As Collection) As InstrumentErrorsEnum

        Dim cmd As String = ""
        Dim temp_collection As New Collection
        Dim item_num, y, temp, count As Integer
        Dim ItemAndValue As ItemClass
        Dim temp_items As New Collection
        Dim temp_counter As Integer
        Dim myItems As New Collection
        Dim x As Integer = 1


        comm.DiscardInBuffer()
        ItemToDownload = Items.Count

        For Each item As String In Items
            myItems.Add(item, item)
        Next

        If myItems.Count < 15 Then
            y = Math.Floor(myItems.Count / 15)
            y = y * 15
        End If

        CommState = CommStateEnum.ReadingGroup

        Do While y < Math.Ceiling(myItems.Count / 15)

            cmd = "RG" & Chr(CommCharEnum.STX)

            'This will give the number of times 15 can go into whatever the size of items is with no remainder
            temp = Math.Floor(myItems.Count / 15)

            'If y - which is a count for how many times we've done this loop - is smaller then our quotient
            'we are safe to use 15 in the loop
            If y < temp Then
                count = 15
            ElseIf y = temp Then
                count = myItems.Count - (y * 15)
            End If

            count += x
            temp_items.Clear()

            Do While x < count
                cmd = cmd & CStr(myItems(x)).PadLeft(3, "0") & ","
                temp_items.Add(myItems(x))
                x += 1
                item_num += 1
            Loop
            cmd = Left(cmd, cmd.Length - 1)

            cmd = cmd & Chr(CommCharEnum.ETX)
            SendDataToComm(cmd)

            temp_counter = 1

            For Each rg_item As Object In ParseRG(CommBuffer)
                Console.Write("Read:" & rg_item & vbNewLine)
                If rg_item <> "! Unsupported" Then
                    ItemAndValue = New ItemClass
                    ItemAndValue.item = temp_items.Item(temp_counter)
                    ItemAndValue.value = rg_item
                    Downloaded.Add(ItemAndValue, temp_items.Item(temp_counter))
                End If
                temp_counter += 1
            Next
            RaiseEvent CurrentItemProgress(item_num)
            y += 1
        Loop

        CommState = CommStateEnum.LinkedIdle

    End Function

    Public Function LR(ByVal ItemNumber As Integer) As Object
        Dim cmd As String
        Dim i_Object As Object

        Try
            CommState = CommStateEnum.ReadingItem
            cmd = "LR" & Chr(CommCharEnum.STX) & ItemNumber.ToString.PadLeft(3, "0") & Chr(CommCharEnum.ETX)
            SendDataToComm(cmd)
            i_Object = ParseRD()
            CommState = CommStateEnum.LinkedIdle
            Return i_Object
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
#End Region

#Region "Protected"
    'This function makes sure everything is in the right condition and opens the comm port
    'Always open the comm port with this function 
    Protected Sub OpenCommPort()
        Try
            If Not comm Is Nothing Then
                If Not comm.IsOpen Then
                    With comm
                        .PortName = Me.Port
                        .BaudRate = Me.BaudRate
                        .NewLine = "\\"
                        .ReadTimeout = 200
                        .WriteTimeout = 150
                        'This will throw an exception if the port is already in use
                        Try
                            .Open()
                        Catch ex As Exception
                            Throw New CommInUseException(comm.PortName)
                        End Try
                        CommState = CommStateEnum.UnlinkedIdle
                    End With
                End If
            Else
                comm = New System.IO.Ports.SerialPort
                OpenCommPort()
            End If
        Catch ex As CommInUseException
            If MessageBox.Show(ex.Message & " Try Again?", "Serial Port Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error) = DialogResult.Retry Then
                Me.OpenCommPort()
            Else
                Throw New Exception("Connection Cancelled.", ex)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    'We can call this function whenever we're expecting a acknowledgement from the instrument
    'This will return NoError if we receive 00 from the instrument or a simple [ACK] in the case of when were Signing on to the instrument
    Protected Function GetReturnCode() As InstrumentErrorsEnum
        Return GetReturnCode(CommBuffer)
    End Function

    'Functions such as ParseRD can now call the function if they encounter any errors from the instrument
    Protected Function GetReturnCode(ByVal IncomingString As String) As InstrumentErrorsEnum
        Dim i = 0, startIndex = 0
        Try
            'I was trying to avoid putting this thread to sleep but this function is being executed before the comm object has
            ' a chance to recieve the event and place it into the CommBuffer, putting this thread to sleep for a 1/4 of a second allows the comm buffer to
            ' catch up
            If IncomingString <> Nothing Then

                If IncomingString.Length >= 9 Then
                    Do While i < IncomingString.Length
                        If IncomingString.Chars(i) = Chr(CommCharEnum.SOH) Then
                            IncomingString = IncomingString.Substring(i + 1)
                            i = startIndex
                        ElseIf IncomingString.Chars(i) = Chr(CommCharEnum.ETX) Then
                            IncomingString = IncomingString.Substring(startIndex, i)
                            Console.Write("Return Value: " & IncomingString & vbNewLine)
                        End If
                        i += 1
                    Loop
                    'if value is not an acknoledgement then check for an error
                    InstrumentError = IncomingString
                ElseIf IncomingString.Length = 1 Then
                    If IncomingString = Chr(CommCharEnum.ACK) Then
                        InstrumentError = InstrumentErrorsEnum.NoError
                    End If
                Else
                    InstrumentError = InstrumentErrorsEnum.CheckByteError
                End If

            Else
                InstrumentError = InstrumentErrorsEnum.NoData
            End If
            Return InstrumentError
        Catch ex As Exception
            Console.Write(ex.Message)
        End Try
    End Function

    Protected Sub Wait()

        Dim w_MessageState As MessageStateEnum = MessageState
        Do
            System.Threading.Thread.Sleep(20)
            Application.DoEvents()
        Loop Until w_MessageState <> MessageState

    End Sub

    'Previously I had this as an event handler, but I was encountering many problems with the delay of the events
    'Now whenever information is passed through the SendDataToComm function, we come to this function to listen for a reply
    'I'm not sure if this is the right way to do this, I havent had any problems with it up to now, but an event might be more robust

    'We don't return any values, we just save the information to the CommBuffer object and then any functions may access it as they wish
    'Be careful when debugging the application, checking the value of the CommBuffer actually removes it from the stack.
    Protected Overridable Sub RcvDataFromComm()
        Dim IncomingData As Object = ""

        Try
            MessageState = MessageStateEnum.WaitingForReceiveEvent
            System.Threading.Thread.Sleep(200)

            Do Until comm.BytesToRead = 0
                Console.Write("Bytes in Comm Buffer: " & comm.BytesToRead & vbNewLine)
                IncomingData += comm.ReadExisting
                System.Threading.Thread.Sleep(50)
            Loop

            If IncomingData <> Nothing Then
                MessageState = MessageStateEnum.ReceivedAPacket
                Console.Write(IncomingData & vbNewLine)
                CommBuffer = IncomingData
            Else
                Throw New NoDataRecievedException("No data received.")
            End If

            MessageState = MessageStateEnum.OK_Idle
        Catch ex As TimeoutException
            Throw New CommExceptions("Read Operation Timedout.")
        Catch ex As NoDataRecievedException
            'We can recall this function again if there was no data received, but I can see this causing an infinite loop
            RcvDataFromComm()
            'Throw New Exception(ex.Message, ex)
        Catch ex As Exception
            Console.Write(ex.Message)
        End Try

    End Sub

    Protected Function CalcCRC(ByVal body As String) As String
        Dim x As Integer
        Dim y As Integer
        Dim InChar As Long
        Dim CRC As Long = 0
        Try
            For x = 1 To Len(body)
                InChar = Asc(Mid$(body, x, 1))
                CRC = CRC Xor (InChar * &H100)
                For y = 1 To 8
                    If CRC And 32768 Then
                        CRC = (CRC * 2) Xor &H1021
                    Else
                        CRC = CRC * 2
                    End If
                Next y
                CRC = CRC And 65535
            Next x
            Return Left$("0000", 4 - Len(Hex(CRC))) & Hex(CRC)
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
            CalcCRC = ""
        End Try
    End Function

    'Any data being sent out to the instrument must go through this method
    'This allows for error correction and the proper flow, logging may be added at a later date as well
    Protected Sub SendDataToComm(ByVal body As String)
        Try
            If comm.IsOpen Then
                If Me.CommState <> CommStateEnum.UnlinkedIdle And MessageState = MessageStateEnum.OK_Idle Then
                    comm.DiscardInBuffer()
                    body = Chr(CommCharEnum.SOH) & body & CalcCRC(body) & Chr(CommCharEnum.EOT)
                    Console.Write("Outgoing >> " & body & vbNewLine)
                    comm.Write(body)
                    RcvDataFromComm()

                Else
                    Me.Connect()
                    SendDataToComm(body)
                End If
            Else
                OpenCommPort()
                SendDataToComm(body)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try

    End Sub

    Protected Sub ClearCommBuffer()
        comm_buffer.Clear()
    End Sub


    'Clear all alarms for the instrument
    'First read the group of alarms and then reset them based on those that are set off
    Protected Function ResetAlarms(ByVal Alarms As Collection) As InstrumentErrorsEnum

        Dim Items As New Collection
        Dim Downloaded As New Collection
        Dim item As ItemClass
        Dim x As Integer = 1

        Items = Alarms

        RG(Items, Downloaded)
        CommState = CommStateEnum.SendingClearAlarmsMessage

        For Each item In Downloaded
            If item.value <> AlarmValuesEnum.NoAlarm Then
                WD(item.item, AlarmValuesEnum.NoAlarm)
                If InstrumentError <> InstrumentErrorsEnum.NoError Then
                    Throw New InstrumentCommunicationException(InstrumentError, item.item)
                End If
            End If
            x += 1
        Next

        CommState = CommStateEnum.LinkedIdle
    End Function

    'A function can call this and it'll use the current value in the comm buffer
    Protected Function ParseRD() As Object
        Return ParseRD(CommBuffer)
    End Function

    'Right now this pulls out the value returned in the string which is passed in
    'It does not error checking on the CRC value though
    'I will add this at a later date, its not a crucial function right now
    Protected Function ParseRD(ByVal IncomingString As String) As Object
        Dim i As Integer = 0, startindex As Integer = 0

        'If instrument returned a value then seperate it
        If IncomingString.Length = 19 Then
            Do While i < IncomingString.Length
                If IncomingString.Chars(i) = Chr(2) Then
                    IncomingString = IncomingString.Substring(i + 1)
                    i = startindex
                ElseIf IncomingString.Chars(i) = Chr(3) Then
                    IncomingString = IncomingString.Substring(startindex, i)
                End If
                i += 1
            Loop
            Return IncomingString
        ElseIf IncomingString.Length > 19 Then
            Return ParseRG(IncomingString)
        ElseIf IncomingString.Length = 9 Then
            'If instrument returned an error code then send it to GetReturnCode
            Return GetReturnCode(IncomingString)
        Else
            Throw New CommExceptions("Incorrect data length." & vbNewLine)
        End If
    End Function


    Protected Function ParseRG()
        Return ParseRG(CommBuffer)
    End Function

    'Pulls out all the values returned by the instrument, and returns a collection of objects
    Protected Function ParseRG(ByVal IncomingString As String) As Collection
        Dim i As Integer = 0, startIndex As Integer = 0
        Dim itemValues As New Collection
        Dim index As Integer
        Dim done As Boolean

        'If our data does not contain any "," then we shouldnt treat it as a group response
        If IncomingString.LastIndexOf(",") <> -1 Then
            IncomingString = Right(IncomingString, IncomingString.Length - 1)
            'Loop if the serial string contains any "," if not exit loop
            Do While done = False
                If IncomingString.Length > 9 Then
                    If IncomingString.IndexOf("!") <> 1 Then
                        If IncomingString.IndexOf(",") <> -1 Then
                            itemValues.Add(Mid(IncomingString, 1, IncomingString.IndexOf(",")))
                            IncomingString = Right(IncomingString, IncomingString.Length - (IncomingString.IndexOf(",") + 1))
                            If IncomingString.Chars(0) = Chr(CommCharEnum.ETX) Then
                                done = True
                            End If
                        Else
                            itemValues.Add(Mid(IncomingString, 1, IncomingString.IndexOf(Chr(CommCharEnum.ETX))))
                            IncomingString = Right(IncomingString, IncomingString.Length - (IncomingString.IndexOf(Chr(3)) + 1))
                        End If
                    Else
                        'Remove "! Unsupported"
                        itemValues.Add(Mid(IncomingString, 2, 13))
                        IncomingString = Right(IncomingString, IncomingString.Length - 16)
                    End If
                Else
                    done = True
                End If
                index = index + 1
            Loop

            Return itemValues
        ElseIf IncomingString.Length = 19 Then
            'If it equals 19 then we received only one item back, sending the data to ParseRD
            Return ParseRD(IncomingString)
        Else
            Throw New Exception("Wrong data length")
        End If
    End Function

#End Region

#End Region

End Class


#Region "Exceptions"

'This is the base class for all Comm Excpetions
Public Class CommExceptions
    Inherits Exception

    Sub New()

    End Sub

    Public Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub
    Public Sub New(ByVal message As String, ByVal inner As Exception)
        MyBase.New(message, inner)
    End Sub

End Class

Public Class CommInUseException
    Inherits CommExceptions

    Private c_PortNumber As String
    Sub New(ByVal CommPort As String)
        MyBase.New("'" & CommPort & "' is currently in use.")
        c_PortNumber = CommPort
    End Sub

    Public ReadOnly Property PortNumber() As String
        Get
            Return c_PortNumber
        End Get
    End Property
End Class

Public Class InstrumentBusyException
    Inherits CommExceptions

    Private e_CommState As CommStateEnum

    Sub New(ByVal CommState As CommStateEnum)
        MyBase.New("Instrument is currently " & CommState.ToString)
        e_CommState = CommState
    End Sub

    Public ReadOnly Property State() As CommStateEnum
        Get
            Return e_CommState
        End Get
    End Property
End Class

Public Class NoPhysicalConnectionException
    Inherits CommExceptions

    Sub New()
        MyBase.New("No physical connection to instrument.")
    End Sub
End Class

#Region "Instrument Errors"

Public Class InstrumentCommunicationException
    Inherits CommExceptions

    Dim i_ItemNumber = Nothing

    Sub New(ByVal CommError As InstrumentErrorsEnum, Optional ByVal ItemNumber As Integer = 0)

        Select Case CommError
            Case InstrumentErrorsEnum.SignOnError
                Throw New SignOnErrorException
            Case InstrumentErrorsEnum.CheckByteError
                Throw New CheckByteException
            Case InstrumentErrorsEnum.FramingError
                Throw New FramingErrorException
            Case InstrumentErrorsEnum.InsufficientPower
                Throw New InsufficientPowerException
            Case InstrumentErrorsEnum.InvalidItemNumberError
                Throw New InvalidItemNumberException(ItemNumber)
            Case InstrumentErrorsEnum.MessageFormatError
                Throw New MessageFormatException
            Case InstrumentErrorsEnum.NoData
                Throw New NoDataRecievedException("No data returned from instrument.")
        End Select
    End Sub
End Class

Public Class SignOnErrorException
    Inherits CommExceptions

    Sub New()
        MyBase.New("Sign On Error.")

    End Sub
End Class
Public Class CheckByteException
    Inherits CommExceptions

    Sub New()
        MyBase.New("CheckByte Error.")
    End Sub
End Class

Public Class MessageFormatException
    Inherits CommExceptions

    Sub New()
        MyBase.New("Message Format Error.")
    End Sub
End Class

Public Class FramingErrorException
    Inherits CommExceptions

    Sub New()
        MyBase.New("Framing Error.")
    End Sub

End Class

Public Class InsufficientPowerException
    Inherits CommExceptions

    Sub New()
        MyBase.New("Insufficient Power for Communication.")
    End Sub
End Class

Public Class InvalidItemNumberException
    Inherits CommExceptions

    Dim item As Integer

    Sub New(ByVal ItemNumber As Integer)
        MyBase.New("Invalid Item Number. " & vbNewLine _
            & "Item " & ItemNumber & " is not accessible.")
        item = ItemNumber
    End Sub

    ReadOnly Property ItemNumber()
        Get
            Return item
        End Get
    End Property
End Class

Public Class NoDataRecievedException
    Inherits CommExceptions

    Sub New(ByVal Message As String)
        MyBase.New(Message)
    End Sub
End Class

#End Region

#End Region

