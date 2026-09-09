'------------------------------------------------------------------------------
' ThermocoupleExampleInAmp.vb
'
' Demonstrates thermocouple configuration and measurement using the
' LJTick-InAmp (commonly used with the T4). If you are using a device such as
' the T7 or T8 which supports the thermocouple AIN_EF, see
' 'ThermocoupleExampleAINEF' instead.
'
' support@labjack.com
'------------------------------------------------------------------------------
Option Explicit On

Imports LabJack


Module ThermocoupleExampleInAmp

    Sub Main()
        Dim handle As Integer
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

            ' Read a thermocouple on AIN6
            Dim channelName As String = "AIN6"
            ' Type K thermocouple reading
            Dim tcType As Integer = LJM.CONSTANTS.ttK
            ' Assuming an InAmp with x51 gain and 1.25V offset is used
            Dim inAmpOffset As Double = 1.25
            Dim inAmpGain As Double = 51
            ' Use the internal temp sensor for CJC
            Dim cjcName As String = "TEMPERATURE_DEVICE_K"
            ' TEMPERATURE_DEVICE_K can be used directly for CJC temperature
            Dim cjcSlope As Integer = 1
            Dim cjcOffset As Integer = 0

            Dim aNames() As String = {channelName, cjcName}
            Dim numFrames As Integer = aNames.Length
            Dim aValues(numFrames) As Double

            ' Read the InAmp output voltage (connected to AIN#) and
            ' the CJC sensor reading.
            LJM.eReadNames(handle, numFrames, aNames, aValues, errorAddress)
            ' Convert the InAmp voltage to the raw thermocouple voltage.
            tcVolts = (aValues(0) - inAmpOffset) / inAmpGain
            ' Apply scaling to the CJC reading if necessary.
            ' At this point, the reading must be in units Kelvin.
            cjTemp = aValues(1) * cjcSlope + cjcOffset
            ' Convert voltage reading to the thermocouple temperature.
            LJM.TCVoltsToTemp(tcType, tcVolts, cjTemp, tcTemp)

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

