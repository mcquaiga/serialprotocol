Public Class MiniMaxClass : Inherits miSerialProtocolClass
    Public Enum BaudRateEnum
        'Mini Max only supports the folllowing baud rates
        b300 = 300
        b600 = 600
        b1200 = 1200
        b2400 = 2400
        b4800 = 4800
        b9600 = 9600
        b19200 = 19200
        b38400 = 38400
    End Enum

    Public Enum AlarmsEnum
        BatteryLowAlarm = 99
        IndexSW1Alarm = 102
        IndexSW2Alarm = 103
        ADAlarm = 104
        AlarmOuput = 108
        PressureLowAlarm = 143
        TemperatureLowAlarm = 144
        PressureHighAlarm = 145
        TemperatureHighAlarm = 146
        DailyCorVolAlarm = 222
        REIAlarm = 435
        RBXAlarm = 176
        UnSupported = 336
    End Enum

    Sub New(ByVal PortNumber As Integer, ByVal BaudRate As BaudRateEnum)
        MyBase.New(PortNumber, BaudRate)
        Me.Instrument = InstrumentTypeCode.MiniMax
    End Sub

    Public Shadows Sub Connect()
        Try
            MyBase.Connect()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Shadows Function ResetAlarms() As InstrumentErrorsEnum
        Dim myAlarms As Array
        Dim alarm As String
        Dim myCollection As New Collection

        myAlarms = System.Enum.GetValues(GetType(AlarmsEnum))

        For Each alarm In myAlarms
            myCollection.Add(CInt(alarm))
        Next

        Return MyBase.ResetAlarms(myCollection)
    End Function

End Class
