'------------------------------------------------------------------------------
' ThermocoupleExampleAINEF.vb
'
' Demonstrates thermocouple configuration and measurement using the
' thermocouple AIN_EF features. This example is for devices such as
' the T7 and T8 that support built-in amplification and the thermocouple
' AIN_EF. For devices without built-in amplification, such as the T4,
' see 'ThermocoupleExampleInAmp' instead.
'
' support@labjack.com
'------------------------------------------------------------------------------
Option Explicit On

Imports LabJack


Module ThermocoupleExampleAINEF
    Public Sub SetupAIN_EF(ByVal handle As Integer, _
        ByVal ainNumber As Integer, _
        ByVal tcType As Integer, _
        ByVal tempUnits As Integer, _
        ByVal cjcAddress As Integer, _
        ByVal cjcSlope As Integer, _
        ByVal cjcOffset As Integer)

        Dim errorAddress As Integer

        Try
            ' For converting LJM TC type constant to TC AIN_EF index.
            ' Thermocouple type:           B   E   J   K   N   R   S   T   C
            Dim tcIndexLut() As Integer = {28, 20, 21, 22, 27, 23, 25, 24, 30}
            Const NUM_FRAMES As Integer = 5
            Dim aNames(NUM_FRAMES) As String
            Dim aValues(NUM_FRAMES) As Double

            ' For setting up the AIN#_EF_INDEX (thermocouple type)
            aNames(0) = "AIN" & ainNumber.ToString() & "_EF_INDEX"
            aValues(0) = tcIndexLut(tcType - 6001)

            ' For setting up the AIN#_EF_CONFIG_A (temperature units)
            aNames(1) = "AIN" & ainNumber.ToString() & "_EF_CONFIG_A"
            aValues(1) = tempUnits

            ' For setting up the AIN#_EF_CONFIG_B (CJC address)
            aNames(2) = "AIN" & ainNumber.ToString() & "_EF_CONFIG_B"
            aValues(2) = cjcAddress

            ' For setting up the AIN#_EF_CONFIG_D (CJC slope)
            aNames(3) = "AIN" & ainNumber.ToString() & "_EF_CONFIG_D"
            aValues(3) = cjcSlope

            ' For setting up the AIN#_EF_CONFIG_E (CJC offset)
            aNames(4) = "AIN" & ainNumber.ToString() & "_EF_CONFIG_E"
            aValues(4) = cjcOffset

            LJM.eWriteNames(handle, NUM_FRAMES, aNames, aValues, errorAddress)

        Catch ljme As LJM.LJMException
            showErrorMessage(ljme)
        End Try
    End Sub

    Public Sub GetReadingsAIN_EF( _
        ByVal handle As Integer, _
        ByVal ainNumber As Integer, _
        ByRef tcTemp As Double, _
        ByRef tcVolts As Double, _
        ByRef cjTemp As Double)

        Dim errorAddress As Integer

        Try
            Dim aNames() As String = { _
                    "AIN" & ainNumber.ToString() & "_EF_READ_A", _
                    "AIN" & ainNumber.ToString() & "_EF_READ_B", _
                    "AIN" & ainNumber.ToString() & "_EF_READ_C"}
            Dim numFrames As Integer = aNames.Length
            Dim aValues(numFrames) As Double
            LJM.eReadNames(handle, numFrames, aNames, aValues, errorAddress)
            tcTemp = aValues(0)   ' Read value from AIN#_EF_READ_A
            tcVolts = aValues(1)  ' Read value from AIN#_EF_READ_B
            cjTemp = aValues(2)   ' Read value from AIN#_EF_READ_C

        Catch ljme As LJM.LJMException
            showErrorMessage(ljme)
        End Try
    End Sub

    Sub Main()
        Dim handle As Integer
        Dim devType As Integer
        Dim cjcAddress As Integer
        Dim tcTemp As Double
        Dim tcVolts As Double
        Dim cjTemp As Double
        Dim errorAddress As Integer

        Try
            ' Open first found LabJack
            LJM.OpenS("ANY", "ANY", "ANY", handle)  ' Any device, Any connection, Any identifier
            'LJM.OpenS("T8", "ANY", "ANY", handle)  ' T8 device, Any connection, Any identifier
            'LJM.OpenS("T7", "ANY", "ANY", handle)  ' T7 device, Any connection, Any identifier
            'LJM.OpenS("T4", "ANY", "ANY", handle)  ' T4 device, Any connection, Any identifier
            'LJM.Open(LJM.CONSTANTS.dtANY, LJM.CONSTANTS.ctANY, "ANY", handle)  ' Any device, Any connection, Any identifier

            displayHandleInfo(handle)
            devType = getDeviceType(handle)

            ' Read a thermocouple on AIN0
            Dim ainNumber As Integer = 0

            ' Type K thermocouple reading
            Dim tcType As Integer = LJM.CONSTANTS.ttK
            Dim tempUnits = 0 '0=K, 1=°C, 2=°F

            ' Use the internal temp sensor for CJC
            ' Note that register names can also be converted to their addresses
            ' using LJM.NameToAddress
            If devType = LJM.CONSTANTS.dtT8 Then
                cjcAddress = 600 + 2 * ainNumber ' TEMPERATURE#
            Else
                cjcAddress = 60052 ' TEMPERATURE_DEVICE_K
            End If
            Dim cjcSlope As Integer = 1 ' 55.56 for LM34
            Dim cjcOffset As Integer = 0 ' 255.37 for LM34

            ' Setup the thermocouple AIN_EF
            SetupAIN_EF(handle, _
                ainNumber, _
                tcType, _
                tempUnits, _
                cjcAddress, _
                cjcSlope, _
                cjcOffset)

            ' Configure an AIN range that can measure up to around 0.07 V and
            ' default resolution. This will use the ±0.1 V setting on the T7
            ' and ± 0.075 V setting on the T8. This can easily be extended for
            ' other configs like AIN#_NEGATIVE_CH for the T7.
            Dim aNames() As String = {"AIN" & ainNumber.ToString & "_RANGE", _
                "AIN" & ainNumber.ToString & "_RESOLUTION_INDEX"}
            Dim aValues() As Double = {0.07, 0}
            LJM.eWriteNames(handle, _
                aNames.Length, _
                aNames, _
                aValues, _
                errorAddress)

            ' Read the thermocouple AIN_EF values
            GetReadingsAIN_EF(handle, ainNumber, tcTemp, tcVolts, cjTemp)
            Console.WriteLine("")
            Console.WriteLine("tcTemp: " & tcTemp.ToString("F3") & " K")
            Console.WriteLine("cjTemp: " & cjTemp.ToString("F3") & " K")
            Console.WriteLine("tcVolts: " & tcVolts.ToString("F6") & " V")

        Catch ljme As LJM.LJMException
            showErrorMessage(ljme)
        End Try

        LJM.Close(handle)  ' Close the device connection

        Console.WriteLine("")
        Console.WriteLine("Done.")
        Console.WriteLine("Press the enter key to exit.")
        Console.ReadLine()  ' Pause for user
    End Sub

End Module

