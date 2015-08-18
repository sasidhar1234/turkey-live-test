Public Class OLDsurvsim
    Public noiterations As Long
    Public outputstringname As String
    Dim myexcel As Microsoft.Office.Interop.Excel.Application
    Dim myworkbook As Microsoft.Office.Interop.Excel.Workbook
    Dim myworksheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim myconnector As STATCONNECTORSRVLib.StatConnector
    Dim mymortsimulator As Mortalitydatasimulator
    Dim diseasemortalityarray(,) As Object
    Dim infectiousatperiodarray(,) As Object
    Dim flocksizearray(,) As Object
    Dim ndaystorun As Long
    Dim tempdailymortalityarr() As Double
    Dim tempdiseasemortalityarr() As Double
    Dim mortfraclimt As Double = 0.002
    Dim targetswabs As Long = 5
    Dim sensitvity As Double = 0.865
    Dim pcrcounter As Long
    Dim detectdayarray(,) As Double
    Dim movementdayarray(,) As Double
    Dim infectiousattesting(,) As Double
    Dim infectiousatmovement(,) As Double
    Dim detectedbeforemovement As Boolean
    Dim movement_day_lower As Long = 3
    Dim movement_day_upper As Long = 6
    Dim randommovementday As Long



    Public Sub New(ByVal con As STATCONNECTORSRVLib.StatConnector)
        myconnector = con
        noiterations = Form1.niterations
    End Sub
    Public Sub initate_excel()
        Try
            myexcel = New Microsoft.Office.Interop.Excel.Application
            myworkbook = myexcel.Workbooks.Open("C:\Users\Sasi\Documents\Work\Turkey transmission\Modeling of Pensylvania\pensylvania\AI Output.xlsx")
        Catch ex As Exception
            myexcel.Quit()
            MsgBox(ex)
        End Try
    End Sub
    Public Sub initiatemortalitysim()

        mymortsimulator = New Mortalitydatasimulator(15, myconnector)
        mymortsimulator.setuprarrays()


    End Sub
    Public Sub getnormalmortality(ByVal curr_iter)
        Dim mymortdata() As Double
        mymortsimulator.getcontinousmortdata(flocksizearray(curr_iter, 1), mymortdata)
        tempdailymortalityarr = mymortdata
    End Sub

    Public Sub initiatearrays()
        myworksheet = myworkbook.Worksheets("Rem")
        diseasemortalityarray = myworksheet.Range(myworksheet.Cells(3, 68), myworksheet.Cells(3 + noiterations - 1, 81)).Value
        flocksizearray = myworksheet.Range(myworksheet.Cells(3, 88), myworksheet.Cells(3 + noiterations - 1, 88)).Value
        myworksheet = myworkbook.Worksheets("Infectious")
        infectiousatperiodarray = myworksheet.Range(myworksheet.Cells(2, 2), myworksheet.Cells(2 + noiterations - 1, 61)).Value



        'flocksizearray = 

        ReDim detectdayarray(noiterations, 0)
        ReDim movementdayarray(noiterations, 0)
        ReDim infectiousatmovement(noiterations, 0)
        ReDim infectiousattesting(noiterations, 0)


    End Sub
    Public Sub get_temp_disease_mort(ByVal curr_iter As Long)
        Dim k As Integer
        ReDim tempdiseasemortalityarr(15)
        For k = 1 To 14
            tempdiseasemortalityarr(k) = diseasemortalityarray(curr_iter, k)
        Next
    End Sub

    Private Function performonedaysurveillance(ByVal Normal As Double, ByVal disease As Integer, ByVal numberoftubes As Integer, ByVal mordetect As Boolean, ByVal mordetectdis As Integer, ByVal everybody As Boolean, ByVal N_pcr_on_day As Long, ByVal flockssize As Long) As Boolean
        Dim Snormal As Integer
        Dim sdisease As Integer
        Dim stotal As Integer
        Dim snumberoftubes As Integer
        Dim snumberdintube As Integer
        Dim pcrpositive As Integer
        Dim snswabs As Integer
        Dim number_swabs As Integer
        Dim n_norm_in_tube As Integer
        Dim detected As Boolean
        Dim myfirstn As Integer
        Dim mysecondn, mythirdn As Integer
        detected = False
        pcrpositive = 0
        'normal mortality
        Snormal = Normal
        sdisease = disease

        If N_pcr_on_day = 2 Then
            If Normal + disease >= 10 Then
                Dim i As Integer
                For i = 1 To 2
                    snumberdintube = hypergeometric(Snormal, sdisease, targetswabs)
                    n_norm_in_tube = targetswabs - snumberdintube
                    If snumberdintube > 0 Then
                        pcrpositive = sim_Binomial(1, sensitvity)
                    End If
                    pcrcounter = pcrcounter + 1
                    Snormal = Snormal - (n_norm_in_tube)
                    If Snormal < 0 Then
                        Snormal = 0
                    End If
                    sdisease = sdisease - (snumberdintube)
                    If pcrpositive > 0 Then
                        detected = True
                    End If
                Next
            Else
                'first tube tested

                myfirstn = Int((Normal + disease) / 2)
                snumberdintube = hypergeometric(Snormal, sdisease, myfirstn)
                n_norm_in_tube = myfirstn - snumberdintube
                If snumberdintube > 0 Then
                    pcrpositive = sim_Binomial(1, sensitvity)
                End If
                pcrcounter = pcrcounter + 1
                Snormal = Snormal - (n_norm_in_tube)
                If Snormal < 0 Then
                    Snormal = 0
                End If
                sdisease = sdisease - (snumberdintube)
                If pcrpositive > 0 Then
                    detected = True
                End If

                'second tube tested
                If Snormal + sdisease >= targetswabs Then
                    mysecondn = targetswabs
                Else
                    mysecondn = Snormal + sdisease
                End If

                snumberdintube = hypergeometric(Snormal, sdisease, mysecondn)
                n_norm_in_tube = mysecondn - snumberdintube
                If snumberdintube > 0 Then
                    pcrpositive = sim_Binomial(1, sensitvity)
                End If
                pcrcounter = pcrcounter + 1
                Snormal = Snormal - (n_norm_in_tube)
                If Snormal < 0 Then
                    Snormal = 0
                End If
                sdisease = sdisease - (snumberdintube)
                If pcrpositive > 0 Then
                    detected = True
                End If

            End If
        End If
        '''''''perform one test********************************************^^^^^^^^^^^^^^^^^^^^^^^$$$$$$$$$$$$##########################
        If N_pcr_on_day = 1 Then
            If Snormal + sdisease >= targetswabs Then
                myfirstn = targetswabs
            Else
                myfirstn = Snormal + sdisease
            End If

            snumberdintube = hypergeometric(Snormal, sdisease, myfirstn)
            n_norm_in_tube = myfirstn - snumberdintube
            If snumberdintube > 0 Then
                pcrpositive = sim_Binomial(1, sensitvity)
            End If
            pcrcounter = pcrcounter + 1
            Snormal = Snormal - (n_norm_in_tube)
            If Snormal < 0 Then
                Snormal = 0
            End If
            sdisease = sdisease - (snumberdintube)
            If pcrpositive > 0 Then
                detected = True
            End If

        End If

        ' peform 3 tests*************************************************************************************************

        If N_pcr_on_day = 3 Then
            If Normal + disease >= 15 Then
                Dim i As Integer
                For i = 1 To 3
                    snumberdintube = hypergeometric(Snormal, sdisease, targetswabs)
                    n_norm_in_tube = targetswabs - snumberdintube
                    If snumberdintube > 0 Then
                        pcrpositive = sim_Binomial(1, sensitvity)
                    End If
                    pcrcounter = pcrcounter + 1
                    Snormal = Snormal - (n_norm_in_tube)
                    If Snormal < 0 Then
                        Snormal = 0
                    End If
                    sdisease = sdisease - (snumberdintube)
                    If pcrpositive > 0 Then
                        detected = True
                    End If
                Next
            Else
                'first tube tested

                myfirstn = Int((Normal + disease) / 3)
                snumberdintube = hypergeometric(Snormal, sdisease, myfirstn)
                n_norm_in_tube = myfirstn - snumberdintube
                If snumberdintube > 0 Then
                    pcrpositive = sim_Binomial(1, sensitvity)
                End If
                pcrcounter = pcrcounter + 1
                Snormal = Snormal - (n_norm_in_tube)
                If Snormal < 0 Then
                    Snormal = 0
                End If
                sdisease = sdisease - (snumberdintube)
                If pcrpositive > 0 Then
                    detected = True
                End If

                'second tube tested

                mysecondn = Int((Snormal + sdisease) / 2)

                snumberdintube = hypergeometric(Snormal, sdisease, mysecondn)
                n_norm_in_tube = mysecondn - snumberdintube
                If snumberdintube > 0 Then
                    pcrpositive = sim_Binomial(1, sensitvity)
                End If
                pcrcounter = pcrcounter + 1
                Snormal = Snormal - (n_norm_in_tube)
                If Snormal < 0 Then
                    Snormal = 0
                End If
                sdisease = sdisease - (snumberdintube)
                If pcrpositive > 0 Then
                    detected = True
                End If

                'third n
                If Snormal + sdisease >= targetswabs Then
                    mythirdn = targetswabs
                Else
                    mythirdn = Snormal + sdisease
                End If

                snumberdintube = hypergeometric(Snormal, sdisease, mythirdn)
                n_norm_in_tube = mythirdn - snumberdintube
                If snumberdintube > 0 Then
                    pcrpositive = sim_Binomial(1, sensitvity)
                End If
                pcrcounter = pcrcounter + 1
                Snormal = Snormal - (n_norm_in_tube)
                If Snormal < 0 Then
                    Snormal = 0
                End If
                sdisease = sdisease - (snumberdintube)
                If pcrpositive > 0 Then
                    detected = True
                End If


            End If
        End If



        'If addsurvlimit = False Then
        If Normal + disease > flockssize * mortfraclimt Then
            mordetect = True
            detected = True
        End If
        'End If
        If detected = True Then
            performonedaysurveillance = True
        End If
    End Function
    Public Function hypergeometric(ByVal normal As Long, ByVal disease As Long, ByVal nswabs As Long)
        Try
            hypergeometric = myconnector.Evaluate("rhyper(1," & disease & "," & normal & "," & nswabs & ")")

        Catch e As Exception
        End Try
    End Function
    Public Function sim_Binomial(ByVal N As Double, ByVal p As Double)
        Dim temp As Object
        temp = myconnector.Evaluate("rbinom(1," & N & "," & p & ")")
        sim_Binomial = temp
    End Function
    Public Sub main()
        noiterations = Form1.niterations
        initate_excel()
        initiatearrays()
        initiatemortalitysim()
        Dim i As Long
        For i = 1 To noiterations
            perform_one_flock_surveillance(i)
        Next

        writeresults()
        myworkbook.Save()
        myexcel.Quit()
        MsgBox("finished")


    End Sub
    Public Function createrandommovementday() As Long
        randommovementday = myconnector.Evaluate("sample(" & movement_day_lower & ":" & movement_day_upper & ", size = 1)")
    End Function
    Public Sub writeresults()
        myworksheet = myworkbook.Worksheets("survpar")
        myworksheet.Range(myworksheet.Cells(3, 6), myworksheet.Cells(3 + noiterations, 6)).Value = detectdayarray
        myworksheet.Range(myworksheet.Cells(3, 7), myworksheet.Cells(3 + noiterations, 7)).Value = movementdayarray
        myworksheet.Range(myworksheet.Cells(3, 8), myworksheet.Cells(3 + noiterations, 8)).Value = infectiousattesting
        myworksheet.Range(myworksheet.Cells(3, 9), myworksheet.Cells(3 + noiterations, 9)).Value = infectiousatmovement
    End Sub

    Public Sub perform_one_flock_surveillance(ByVal curr_iter As Long)
        Dim normmortality As Double
        Dim mordetect As Boolean
        Dim actdayconeggs(15) As Single
        Dim testdayarray(15) As Integer
        Dim totaleggstobemoved As Single
        Dim diseasetimetilltotal As Single
        Dim acttodisease(15) As Integer
        Dim actdetectday As Integer
        Dim actdayindex As Integer
        Dim eggmovestartday As Integer
        Dim eggmoveendday As Integer
        Dim mycase As Integer
        mordetect = True
        getnormalmortality(curr_iter)
        createrandommovementday()
        get_temp_disease_mort(curr_iter)
        Dim tempdifference As Integer


        mycase = 0

        For q = 1 To 15
            testdayarray(q) = 0
        Next q

        If mycase = 0 Then
            testdayarray(randommovementday) = 3
            'If randommovementday - 1 >= 1 Then
            '    testdayarray(randommovementday - 1) = 1
            'End If
            'If randommovementday - 2 >= 1 Then
            '    testdayarray(randommovementday - 2) = 1
            'End If
            'eggmovestartday = 2 - minhldtime + 1
            'If eggmovestartday < 1 Then eggmovestartday = 1
            'eggmoveendday = 6 - minhldtime
        Else

        End If



        Dim detected As Boolean
        detected = False
        Dim disease_detectday As Integer
        disease_detectday = 0

        Dim i As Integer
        For i = 1 To 14

            Dim testpcr As Integer

            testpcr = testdayarray(i)
            'testpcr = False

            Dim mordetdismor As Integer
            mordetdismor = tempdailymortalityarr(i)



            detected = performonedaysurveillance(tempdailymortalityarr(i - 1), tempdiseasemortalityarr(i - 1), 2, True, mordetdismor, False, testpcr, flocksizearray(curr_iter, 1))


            If detected Then
                disease_detectday = i
                actdetectday = i
                Exit For
            End If
        Next

        If detected = False Then
            disease_detectday = 15
        End If


        'For i = eggmovestartday To eggmoveendday
        '    If acttodisease(i) > 0 Then
        '        totaleggstobemoved = totaleggstobemoved + dailyeggproduction(acttodisease(i))
        '    End If
        'Next



        'For i = 1 To 7
        '    If acttodisease(i) > 0 Then
        '        actdayconeggs(i) = dailyeggproduction(acttodisease(i))
        '    End If
        'Next


        'special
        '' If actdetectday < 3 Then
        '' actdayconeggs(1) = 0
        '' End If
        ''

        'Dim eggmax As Double
        'Dim eggsum As Double
        'Dim k As Integer
        'For k = 1 To 7
        '    Sheet6.Cells(2 + curiternumber, 5 + k) = actdayconeggs(k)
        'Next k

        'Sheet6.Cells(2 + curiternumber, 5) = curiternumber
        detectdayarray(curr_iter - 1, 0) = disease_detectday
        movementdayarray(curr_iter - 1, 0) = randommovementday
        infectiousattesting(curr_iter - 1, 0) = infectiousatperiodarray(curr_iter, Math.Round(randommovementday * (24 / Form1.tau)))
        infectiousatmovement(curr_iter - 1, 0) = infectiousatperiodarray(curr_iter, Math.Round(randommovementday * (24 / Form1.tau) + 12 / Form1.tau))

        Beep()


    End Sub
End Class
