Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.IO
Imports ClosedXML.Excel

Public Class _Default
    Inherits Page

    Dim Conn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("db").ConnectionString)
    Private Property cmd As SqlCommand
    Private Property cmd2 As SqlCommand
    Dim rdr As SqlDataReader
    Dim rdr2 As SqlDataReader

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim wb As XLWorkbook
        Dim rowCount As Boolean
        Dim count As Integer
        Dim deviceId As String


        wb = New XLWorkbook(FileUpload1.PostedFile.InputStream)
        Dim ws = wb.Worksheet(1)
        count = 1
        rowCount = True

        Conn.Open()
        While rowCount
            Dim endLine As Boolean
            Dim strEnd As String
            strEnd = ws.Cell(count, 1).GetString
            endLine = InStr(strEnd, "Centlec")
            If endLine = True Then
                rowCount = False
                Exit While
            Else
                'check if rec exist otherwise update
                If count <> 1 Then
                    deviceId = ws.Cell(count, 1).GetString
                    Dim strCheck As String
                    Dim recFound As Boolean
                    strCheck = "select [Device _ID] FROM [Centlec08_v2].[dbo].[MDM_Devices] where [Device _ID] = '" & deviceId & "'"
                    Using cmd As New SqlCommand
                        With cmd
                            .Connection = Conn
                            .CommandText = Data.CommandType.Text
                            .CommandText = strCheck
                        End With
                        rdr = cmd.ExecuteReader
                        If rdr.HasRows Then
                            recFound = True
                        Else
                            recFound = False
                        End If
                        rdr.Close()
                    End Using
                    If recFound Then
                        updateRec(wb, count, "update")
                    Else
                        updateRec(wb, count, "insert")
                    End If
                End If
            End If
            count = count + 1
        End While

        Conn.Close()



    End Sub

    'Private Sub insertRec(wb As XLWorkbook, row As Integer)

    '    Dim insert As String
    '    Dim rowCount As Boolean
    '    Dim errLog As String
    '    Dim count As Integer
    '    Dim mdgrpID As Integer

    '    insert = "insert into [Centlec08_v2].[dbo].[MDM_Devices] ([Device _ID], [Device Seq No], [Configuration Name], [Device Name] , [OS] , [OS Version] , [Model], [IMEI Number] , [IMSI Number] , [ICCID Number] , [Serial Number] , [Mobilock Pro Version] , [MD Group], [Profile Name], [Phone Number], [Last Seen] , [Added On] , [Last Location Address] , [Last Location Latitude] , [Last Location Longitude] , [Last Location Time] , [Status] , [Monthly Data Usage (As Per Date)] , [TimeZone] , [Notes], [Licence Code] , [Wifi Mac Address] , [Wifi Frequency Band] , [Device Firmware] , [Country Code] , [Profile Name1] , [BA_GroupID] " &
    '        "Values(@Device_ID, @DeviceSeqNo, @ConfigurationName, @DeviceName, @os, @OSVersion, @Model, @IMEINumber, @IMSINumber, @ICCIDNumber, @SerialNumber, @MbilockProVersion, @MDGroup, @ProfileName, @PhoneNumber, @LastSeen, @AddedOn, @LastLocationAddress, @LastLocationLatitude, @LastLocationLongitude, @LastLocationTime, @Status, @MonthlyDataUsage, @TimeZone, @Notes, @LicenceCode, @WifiMacAddress, @WifiFrequencyBand, @DeviceFirmware, @CountryCode, @ProfileName1, @BA_GroupID, @DeviceID2)"
    '    Dim ws = wb.Worksheet(1)
    '    count = 1

    '    Dim lastrow = ws.RangeUsed
    '    errLog = ""
    '    rowCount = True

    'End Sub

    Private Sub updateRec(wb As XLWorkbook, row As Integer, recType As String)

        Dim update, insert As String
        Dim errLog As String
        Dim count As Integer
        Dim mdgrpID As Integer

        update = "Update [Centlec08_v2].[dbo].[MDM_Devices] set [Device _ID] = @DeviceID, [Device Seq No] = @DeviceSeqNo, [Configuration Name] = @ConfigurationName, [Device Name] =@DeviceName, [OS] = @os, [OS Version] = @OSVersion, [Model] = @Model, [IMEI Number] = @IMEINumber, [IMSI Number] = @IMSINumber, [ICCID Number] = @ICCIDNumber, [Serial Number] = @SerialNumber, [Mobilock Pro Version] = @MbilockProVersion, [MD Group]=@MDGroup, [Profile Name]=@ProfileName, [Phone Number] = @PhoneNumber, [Last Seen] = @LastSeen, [Added On] = @AddedOn, [Last Location Address] = @LastLocationAddress, [Last Location Latitude] = @LastLocationLatitude, [Last Location Longitude] = @LastLocationLongitude, [Last Location Time] = @LastLocationTime, [Status] = @Status, [Monthly Data Usage (As Per Date)] = @MonthlyDataUsage, [TimeZone] = @TimeZone, [Notes] = @Notes, [Licence Code] = @LicenceCode, [Wifi Mac Address] = @WifiMacAddress, [Wifi Frequency Band] = @WifiFrequencyBand, [Device Firmware] = @DeviceFirmware, [Country Code] = @CountryCode, [Profile Name1] = @ProfileName1, [BA_GroupID] = @BA_GroupID where [Device _ID] =  @DeviceID2"
        insert = "insert into [Centlec08_v2].[dbo].[MDM_Devices] ([Device _ID], [Device Seq No], [Configuration Name], [Device Name] , [OS] , [OS Version] , [Model], [IMEI Number] , [IMSI Number] , [ICCID Number] , [Serial Number] , [Mobilock Pro Version] , [MD Group], [Profile Name], [Phone Number], [Last Seen] , [Added On] , [Last Location Address] , [Last Location Latitude] , [Last Location Longitude] , [Last Location Time] , [Status] , [Monthly Data Usage (As Per Date)] , [TimeZone] , [Notes], [Licence Code] , [Wifi Mac Address] , [Wifi Frequency Band] , [Device Firmware] , [Country Code] , [Profile Name1] , [BA_GroupID], [BA_ConditionID], [BA_StatusID], [BA_Status], [BA_Condition]) " &
            "Values(@Device_ID, @DeviceSeqNo, @ConfigurationName, @DeviceName, @os, @OSVersion, @Model, @IMEINumber, @IMSINumber, @ICCIDNumber, @SerialNumber, @MbilockProVersion, @MDGroup, @ProfileName, @PhoneNumber, @LastSeen, @AddedOn, @LastLocationAddress, @LastLocationLatitude, @LastLocationLongitude, @LastLocationTime, @Status, @MonthlyDataUsage, @TimeZone, @Notes, @LicenceCode, @WifiMacAddress, @WifiFrequencyBand, @DeviceFirmware, @CountryCode, @ProfileName1, @BA_GroupID, @BA_ConditionID, @BA_StatusID, @BA_Status, @BA_Condition)"
        'wb = New XLWorkbook(FileUpload1.PostedFile.InputStream)
        Dim ws = wb.Worksheet(1)
        count = 1

        '' Conn.Open()
        Dim lastrow = ws.RangeUsed
        errLog = ""

        'get group ID
        Dim sel As String
            sel = "SELECT distinct BA_GroupID,[MDGroup] FROM [Centlec08_v2].[dbo].[MDM_Devices] A, [Prod_CentlecBP].[dbo].[MobileDeviceGroup] B WHERE A.[MD Group] = B.[MDGroup]"
        Using cmd2 As New SqlCommand
            With cmd2
                .Connection = Conn
                .CommandText = Data.CommandType.Text
                .CommandText = sel
            End With
            rdr2 = cmd2.ExecuteReader
            If rdr2.HasRows Then
                Do While rdr2.Read
                    If LCase(rdr2.Item("MDGroup")) = LCase(ws.Cell(row, 13).GetString) Then
                        mdgrpID = rdr2.Item("BA_GroupID")
                    End If
                Loop
            End If
            rdr2.Close()
        End Using

        Using cmd2 As New SqlCommand
            Dim devId As String
            devId = ws.Cell(row, 1).GetString
            With cmd2
                .Connection = Conn
                .CommandType = Data.CommandType.Text
                If recType = "update" Then
                    .CommandText = update
                Else
                    .CommandText = insert
                End If
                .Parameters.AddWithValue("@Device_ID", ws.Cell(row, 1).GetString)
                .Parameters.AddWithValue("@DeviceSeqNo", ws.Cell(row, 2).GetString)
                .Parameters.AddWithValue("@ConfigurationName", ws.Cell(row, 3).GetString)
                .Parameters.AddWithValue("@DeviceName", ws.Cell(row, 4).GetString)
                .Parameters.AddWithValue("@os", ws.Cell(row, 5).GetString)
                .Parameters.AddWithValue("@OSVersion", ws.Cell(row, 6).GetString)
                .Parameters.AddWithValue("@Model", ws.Cell(row, 7).GetString)
                .Parameters.AddWithValue("@IMEINumber", ws.Cell(row, 8).GetString)
                .Parameters.AddWithValue("@IMSINumber", ws.Cell(row, 9).GetString)
                .Parameters.AddWithValue("@ICCIDNumber", ws.Cell(row, 10).GetString)
                .Parameters.AddWithValue("@SerialNumber", ws.Cell(row, 11).GetString)
                .Parameters.AddWithValue("@MbilockProVersion", ws.Cell(row, 12).GetString)
                .Parameters.AddWithValue("@MDGroup", ws.Cell(row, 13).GetString)
                .Parameters.AddWithValue("@ProfileName", ws.Cell(row, 14).GetString)
                .Parameters.AddWithValue("@PhoneNumber", ws.Cell(row, 15).GetString)
                .Parameters.AddWithValue("@LastSeen", ws.Cell(row, 16).GetString)
                .Parameters.AddWithValue("@AddedOn", ws.Cell(row, 17).GetString)
                .Parameters.AddWithValue("@LastLocationAddress", ws.Cell(row, 18).GetString)
                .Parameters.AddWithValue("@LastLocationLatitude", ws.Cell(row, 19).GetString)
                .Parameters.AddWithValue("@LastLocationLongitude", ws.Cell(row, 20).GetString)
                .Parameters.AddWithValue("@LastLocationTime", ws.Cell(row, 21).GetString)
                .Parameters.AddWithValue("@Status", ws.Cell(row, 22).GetString)
                .Parameters.AddWithValue("@MonthlyDataUsage", ws.Cell(row, 23).GetString)
                .Parameters.AddWithValue("@TimeZone", ws.Cell(row, 24).GetString)
                .Parameters.AddWithValue("@Notes", ws.Cell(row, 25).GetString)
                .Parameters.AddWithValue("@LicenceCode", ws.Cell(row, 26).GetString)
                .Parameters.AddWithValue("@WifiMacAddress", ws.Cell(row, 27).GetString)
                .Parameters.AddWithValue("@WifiFrequencyBand", ws.Cell(row, 28).GetString)
                .Parameters.AddWithValue("@DeviceFirmware", ws.Cell(row, 29).GetString)
                .Parameters.AddWithValue("@CountryCode", ws.Cell(row, 30).GetString)
                .Parameters.AddWithValue("@ProfileName1", ws.Cell(row, 31).GetString)
                .Parameters.AddWithValue("@BA_GroupID", mdgrpID)
                If recType = "update" Then
                    .Parameters.AddWithValue("@DeviceID2", ws.Cell(row, 1).GetString)
                End If
                If recType = "insert" Then
                    .Parameters.AddWithValue("@BA_ConditionID", 1)
                    .Parameters.AddWithValue("@BA_StatusID", 1)
                    .Parameters.AddWithValue("@BA_Status", "Issued")
                    .Parameters.AddWithValue("@BA_Condition", "Active")
                End If
            End With
            Try
                If devId >= 1 Then
                    cmd2.ExecuteNonQuery()
                    lblResult.Text = " Successful upload: " & "Total Rows: " & row - 1
                    lblResult.Visible = True
                End If
            Catch ex As Exception
                If devId <> "" Then
                    errLog = errLog & vbNewLine & ex.Message
                    lblResult.Text = errLog & " " & "Rowcount: " & row
                    lblResult.Visible = True
                    Exit Sub
                End If

            End Try
        End Using
        '' Conn.Close()

    End Sub
End Class