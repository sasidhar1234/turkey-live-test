Public Class survsim
    Public noiterations As Long
    Public outputstringname As String
    Dim myexcel As Microsoft.Office.Interop.Excel.Application
    Dim myworkbook As Microsoft.Office.Interop.Excel.Workbook
    Dim myworksheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim myconnector As STATCONNECTORSRVLib.StatConnector
    Dim mymortsimulator As Mortalitydatasimulator
    Dim myhenmortsimulator As Mortalityhendatasimulator
    Dim diseasemortalityarray(,) As Object
    Dim infectiousatperiodarray(,) As Object
    Dim egg_productrate_period_array(,) As Object
    'Dim contam_egg_array(,) As Object
    Dim eggratelower As Single = 0.45
    Dim eggratehigher As Single = 0.7
    Dim ndaystorun As Long
    Dim tempdailymortalityarr() As Double
    Dim tempdiseasemortalityarr() As Double
    Dim temp_acmortality As Double
    Dim tempeggratearr() As Double
    Dim flocksizearray(,) As Object
    'Dim temp_diseaseeggproductarr() As Double
    Dim mortfraclimt As Double = 0.003 'adjust to 0.003
    Dim eggdecreaslimt As Double = 0.05
    Dim targetswabs As Long = 11 'PCR testing pool size
    Dim ACtargetswabs As Long = 5 ''antigen capture testing pool size
    Dim disease_Mort_atperiodarray(,) As Object
    Dim clinical_at_egg_periodarray(,) As Object
    Dim sensitvity As Double = 0.865
    Dim acsensitvity As Double = 0.6
    Dim dropinegg As Double = 0.3
    Dim pcrcounter As Long
    Dim detectdayarray(,) As Double
    Dim egg_detectarray(,) As Double
    Dim mort_detectarray(,) As Double
    Dim movementdayarray(,) As Double
    Dim infectiousattesting(,) As Double
    Dim infectiousatmovement(,) As Double
    Dim infe_innot_detct(,) As Object
    Dim not_detect_count As Long
    Dim not_detect_count_withinfectiousbirds As Long
    Dim niterations_with_infectiousbirds As Long

    Dim detectedbeforemovement As Boolean
    'Dim egg_moved_array(,) As Double
    Dim minhldtime As Integer = 2
    Public movement_day_lower As Long = 5 ''minimum days flock infected before movement
    Public movement_day_upper As Long = 10 ''maximum days flock infected before movement
    Dim Disease_day_on_movement_day As Long
    Dim ndaysmovedpershipment As Integer = 5
    Dim specialcommentstring As String
    Dim specialcommentstring_2 As String
    Public isonetubeper50 As Boolean = True
    Dim final_detect_perc As Double
    Dim min_min_detectarray(,) As Double
    Dim is_Ac_test_pooled As Boolean = True
    Dim is_daily_PCRtesting As Boolean = False
    Dim is_daily_ACtesting As Boolean = False
    Dim autoclose As Boolean = True
    Dim nperiodsinaday As Integer
    Dim iterstart As Integer = 1
    Dim iterend As Integer = 6000


    Public Function getaverage(ByVal mydblarray(,) As Object) As Double
        Dim myreturn As Object
        Try
            myconnector.SetSymbol("zmeanz", mydblarray)
            myreturn = myconnector.Evaluate("mean(zmeanz)")
            getaverage = CDbl(myreturn)

        Catch ex As Exception
            getaverage = 0
        End Try

    End Function
    Public Function getquantile(ByVal mydblarray(,) As Object, ByVal tpercentile As Double) As Double
        Dim myreturn As Object
        Try
            myconnector.SetSymbol("zmeanz", mydblarray)
            myreturn = myconnector.Evaluate("quantile(zmeanz," & tpercentile & ")")
            getquantile = CDbl(myreturn)
        Catch ex As Exception
            getquantile = 0
        End Try
    End Function
    Public Sub suvrunparamterrecorder()
        myworksheet = myworkbook.Worksheets("Runp")
        myworksheet.Cells(18, 2) = flocksizearray.GetLength(0)
        If form1.is_turkey_hen = False Then
            myworksheet.Cells(19, 2) = "turkey Toms"
        Else
            myworksheet.Cells(19, 2) = "turkey hens"
        End If
        myworksheet.Cells(21, 2) = getaverage(flocksizearray)
        myworksheet.Cells(22, 2) = Form1.starttime
        myworksheet.Cells(23, 2) = Form1.endtime
        myworksheet.Cells(24, 2) = DateAndTime.DateDiff(DateInterval.Second, Form1.starttime, Form1.endtime)
        myworksheet.Cells(25, 2) = minhldtime
        myworksheet.Cells(26, 2) = movement_day_upper
        myworksheet.Cells(27, 2) = movement_day_lower
        myworksheet.Cells(28, 2) = targetswabs
        myworksheet.Cells(28, 3) = specialcommentstring
        myworksheet.Cells(28, 5) = specialcommentstring_2
    End Sub

    Public Sub New(ByVal con As STATCONNECTORSRVLib.StatConnector)
        myconnector = con
        noiterations = Form1.niterations
    End Sub
    Public Sub initate_excel()
        Try
            myexcel = New Microsoft.Office.Interop.Excel.Application
            myworkbook = myexcel.Workbooks.Open("C:\Codes\Live turkey Movement\AIoutput\AI Output.xlsx")
        Catch ex As Exception
            myexcel.Quit()
            MsgBox(ex)
        End Try
    End Sub

    Public Sub initiate_mortalitysim()
        If form1.is_turkey_hen = False Then
            mymortsimulator = New Mortalitydatasimulator(15, myconnector)
            mymortsimulator.setuprarrays()
        Else
            myhenmortsimulator = New Mortalityhendatasimulator(15, myconnector)
            myhenmortsimulator.setuprarrays()
        End If



    End Sub
    Public Sub getnormalmortality(ByVal curr_iter, ByVal Disease_day_on_movement_day)

        Dim mymortdata() As Double
        If form1.is_turkey_hen = False Then
            mymortsimulator.getcontinousmortdata(flocksizearray(curr_iter, 1), mymortdata)
        Else
            myhenmortsimulator.getcontinousmortdata(flocksizearray(curr_iter, 1), mymortdata)
        End If

        ReDim tempdailymortalityarr(15)
        Dim k As Integer
        For k = 0 To Disease_day_on_movement_day - 1
            'this is complicated but matching the mortality data to the random movement day. So randmovd
            tempdailymortalityarr(k) = mymortdata(13 - (Disease_day_on_movement_day - 1) + k)
        Next
        If Disease_day_on_movement_day < 14 Then
            For k = Disease_day_on_movement_day To 13
                tempdailymortalityarr(k) = mymortdata(13)
            Next
        End If
    End Sub

    Public Sub initiatearrays()
        myworksheet = myworkbook.Worksheets("Rem")
        disease_Mort_atperiodarray = myworksheet.Range(myworksheet.Cells(2 + iterstart - 1, 2), myworksheet.Cells(2 + iterend - 1, 2 + form1.nperiods + 1)).Value
        flocksizearray = myworksheet.Range(myworksheet.Cells(3 + iterstart - 1, 88), myworksheet.Cells(3 + iterend - 1, 88)).Value
        myworksheet = myworkbook.Worksheets("Infectious")
        infectiousatperiodarray = myworksheet.Range(myworksheet.Cells(2 + iterstart - 1, 2), myworksheet.Cells(2 + iterend - 1, 61)).Value
        'myworksheet = myworkbook.Worksheets("clinical infec")
        'contam_egg_array = myworksheet.Range(myworksheet.Cells(2, 65), myworksheet.Cells(2 + noiterations - 1, 78)).Value
        'myworksheet = myworkbook.Worksheets("Egg")
        egg_productrate_period_array = myworksheet.Range(myworksheet.Cells(2 + iterstart - 1, 2), myworksheet.Cells(2 + iterend - 1, 61)).Value
        myworksheet = myworkbook.Worksheets("Flocksize")
        egg_productrate_period_array = myworksheet.Range(myworksheet.Cells(3 + iterstart - 1, 5), myworksheet.Cells(3 + iterend - 1, 5)).Value
        flocksizearray = myworksheet.Range(myworksheet.Cells(3 + iterstart - 1, 2), myworksheet.Cells(3 + iterend - 1, 2)).Value
        myworksheet = myworkbook.Worksheets("Infectious")
        infectiousatperiodarray = myworksheet.Range(myworksheet.Cells(2 + iterstart - 1, 2), myworksheet.Cells(2 + iterend - 1, 2 + form1.nperiods + 1)).Value



        ReDim detectdayarray(iterend - iterstart + 1, 0)
        ReDim movementdayarray(iterend - iterstart + 1, 0)
        ReDim infectiousatmovement(iterend - iterstart + 1, 0)
        ReDim infectiousattesting(iterend - iterstart + 1, 0)
        'ReDim egg_moved_array(noiterations, 0)
        ReDim egg_detectarray(iterend - iterstart + 1, 0)
        ReDim mort_detectarray(iterend - iterstart + 1, 0)
        ReDim min_min_detectarray(iterend - iterstart + 1, 0)
    End Sub
    Public Sub get_temp_disease_mort(ByVal curr_iter As Long, movementday As Long)
        Dim k As Integer
        ReDim tempdiseasemortalityarr(15)
        For k = 1 To 14
            tempdiseasemortalityarr(k) = disease_Mort_atperiodarray(curr_iter, (k * nperiodsinaday + 1)) - disease_Mort_atperiodarray(curr_iter, ((k - 1) * nperiodsinaday + 1))
        Next
    End Sub


    Public Sub get_temp_ac_sp_mortality(ByVal curr_iter As Long, movementday As Long, ByVal movement_integer As Integer)
        Dim tempmovement_integer As Integer
        tempmovement_integer = movement_integer

        'Math.Round((movementdouble * 24) / (24 / form1.tau))

        temp_acmortality = disease_Mort_atperiodarray(curr_iter, (movementday * nperiodsinaday + 1 + tempmovement_integer)) - disease_Mort_atperiodarray(curr_iter, (movementday * nperiodsinaday + 1))
    End Sub



    'Private Function perform_egg_surv(ByVal eggproductrate) As Boolean
    'If eggproductrate <= (Form1.baseline_egg_productionrate * (1 - eggdecreaslimt)) Then perform_egg_surv = True

    ' End Function

    Private Function perform_mort_surv(ByVal Normal As Double, ByVal disease As Integer, ByVal flockssize As Long) As Boolean
        If Normal + disease > flockssize * mortfraclimt Then
            perform_mort_surv = True
        Else
            perform_mort_surv = False
        End If
    End Function



    Private Function performoneACTomdaysurveillance(ByVal Normal As Double, ByVal disease As Integer, ByVal doactest As Long) As Boolean
        Dim Snormal As Integer
        Dim sdisease As Integer
        Dim stotal As Integer
        Dim snumberoftubes As Integer
        Dim snumberdintube As Integer
        Dim ACpositive As Integer
        Dim snswabs As Integer
        Dim number_swabs As Integer
        Dim n_norm_in_tube As Integer
        Dim detected As Boolean
        Dim myfirstn As Integer
        Dim mysecondn, mythirdn As Integer
        Snormal = Normal
        sdisease = disease
        detected = False
        ACpositive = 0
        'normal mortality

        If is_Ac_test_pooled = False Then
            If doactest >= 1 Then
                detected = False
                ACpositive = 0
                'normal mortality

                If sdisease >= ACtargetswabs Then
                    myfirstn = ACtargetswabs
                Else
                    myfirstn = sdisease
                End If

                Dim tempositives As Long


                tempositives = sim_Binomial(myfirstn, acsensitvity)
                If tempositives >= 1 Then
                    detected = True
                End If

                'End If
                If detected = True Then
                    performoneACTomdaysurveillance = True
                End If
            End If
        End If

        ''''New one similar to PCR
        If is_Ac_test_pooled Then
            If doactest = 2 Then
                If Snormal + sdisease >= 2 * ACtargetswabs Then
                    Dim i As Integer
                    For i = 1 To 2
                        snumberdintube = hypergeometric(Snormal, sdisease, ACtargetswabs)
                        n_norm_in_tube = ACtargetswabs - snumberdintube
                        If snumberdintube > 0 Then
                            ACpositive = sim_Binomial(1, acsensitvity)
                        End If
                        pcrcounter = pcrcounter + 1
                        Snormal = Snormal - (n_norm_in_tube)
                        If Snormal < 0 Then
                            Snormal = 0
                        End If
                        sdisease = sdisease - (snumberdintube)
                        If ACpositive > 0 Then
                            detected = True
                        End If
                    Next
                Else
                    'first tube tested

                    myfirstn = Int((Snormal + sdisease) / 2)
                    snumberdintube = hypergeometric(Snormal, sdisease, myfirstn)
                    n_norm_in_tube = myfirstn - snumberdintube
                    If snumberdintube > 0 Then
                        ACpositive = sim_Binomial(1, acsensitvity)
                    End If
                    pcrcounter = pcrcounter + 1
                    Snormal = Snormal - (n_norm_in_tube)
                    If Snormal < 0 Then
                        Snormal = 0
                    End If
                    sdisease = sdisease - (snumberdintube)
                    If ACpositive > 0 Then
                        detected = True
                    End If

                    'second tube tested
                    If Snormal + sdisease >= ACtargetswabs Then
                        mysecondn = ACtargetswabs
                    Else
                        mysecondn = Snormal + sdisease
                    End If

                    snumberdintube = hypergeometric(Snormal, sdisease, mysecondn)
                    n_norm_in_tube = mysecondn - snumberdintube
                    If snumberdintube > 0 Then
                        ACpositive = sim_Binomial(1, acsensitvity)
                    End If
                    Snormal = Snormal - (n_norm_in_tube)
                    If Snormal < 0 Then
                        Snormal = 0
                    End If
                    sdisease = sdisease - (snumberdintube)
                    If ACpositive > 0 Then
                        detected = True
                    End If

                End If
            End If
            '''''''perform one test********************************************^^^^^^^^^^^^^^^^^^^^^^^$$$$$$$$$$$$##########################
            If doactest = 1 Then
                If Snormal + sdisease >= ACtargetswabs Then
                    myfirstn = ACtargetswabs
                Else
                    myfirstn = Snormal + sdisease
                End If

                snumberdintube = hypergeometric(Snormal, sdisease, myfirstn)
                n_norm_in_tube = myfirstn - snumberdintube
                If snumberdintube > 0 Then
                    ACpositive = sim_Binomial(1, acsensitvity)
                End If
                pcrcounter = pcrcounter + 1
                Snormal = Snormal - (n_norm_in_tube)
                If Snormal < 0 Then
                    Snormal = 0
                End If
                sdisease = sdisease - (snumberdintube)
                If ACpositive > 0 Then
                    detected = True
                End If

            End If

            ' peform 3 tests*************************************************************************************************

            If doactest = 3 Then
                If Snormal + sdisease >= 3 * ACtargetswabs Then
                    Dim i As Integer
                    For i = 1 To 3
                        snumberdintube = hypergeometric(Snormal, sdisease, ACtargetswabs)
                        n_norm_in_tube = ACtargetswabs - snumberdintube
                        If snumberdintube > 0 Then
                            ACpositive = sim_Binomial(1, acsensitvity)
                        End If
                        pcrcounter = pcrcounter + 1
                        Snormal = Snormal - (n_norm_in_tube)
                        If Snormal < 0 Then
                            Snormal = 0
                        End If
                        sdisease = sdisease - (snumberdintube)
                        If ACpositive > 0 Then
                            detected = True
                        End If
                    Next
                Else
                    'first tube tested

                    myfirstn = Int((Snormal + sdisease) / 3)
                    snumberdintube = hypergeometric(Snormal, sdisease, myfirstn)
                    n_norm_in_tube = myfirstn - snumberdintube
                    If snumberdintube > 0 Then
                        ACpositive = sim_Binomial(1, acsensitvity)
                    End If
                    pcrcounter = pcrcounter + 1
                    Snormal = Snormal - (n_norm_in_tube)
                    If Snormal < 0 Then
                        Snormal = 0
                    End If
                    sdisease = sdisease - (snumberdintube)
                    If ACpositive > 0 Then
                        detected = True
                    End If

                    'second tube tested

                    mysecondn = Int((Snormal + sdisease) / 2)

                    snumberdintube = hypergeometric(Snormal, sdisease, mysecondn)
                    n_norm_in_tube = mysecondn - snumberdintube
                    If snumberdintube > 0 Then
                        ACpositive = sim_Binomial(1, acsensitvity)
                    End If
                    pcrcounter = pcrcounter + 1
                    Snormal = Snormal - (n_norm_in_tube)
                    If Snormal < 0 Then
                        Snormal = 0
                    End If
                    sdisease = sdisease - (snumberdintube)
                    If ACpositive > 0 Then
                        detected = True
                    End If

                    'third n
                    If Snormal + sdisease >= ACtargetswabs Then
                        mythirdn = ACtargetswabs
                    Else
                        mythirdn = Snormal + sdisease
                    End If

                    snumberdintube = hypergeometric(Snormal, sdisease, mythirdn)
                    n_norm_in_tube = mythirdn - snumberdintube
                    If snumberdintube > 0 Then
                        ACpositive = sim_Binomial(1, acsensitvity)
                    End If
                    pcrcounter = pcrcounter + 1
                    Snormal = Snormal - (n_norm_in_tube)
                    If Snormal < 0 Then
                        Snormal = 0
                    End If
                    sdisease = sdisease - (snumberdintube)
                    If ACpositive > 0 Then
                        detected = True
                    End If
                End If
            End If

        End If



        'End If
        If detected = True Then
            performoneACTomdaysurveillance = True
        End If


    End Function
    Private Function performonedaysurveillance_withlive(ByVal Normal As Double, ByVal disease As Integer, ByVal N_pcr_on_day As Long, Iteration As Long, testperiod As Long) As Boolean
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
        Dim s_infectious As Long
        Dim snbirds_alive As Long
        Dim lprevaldbl As Double
        Dim templbirdstobetested, nposlivebirds_pertube As Double
        s_infectious = infectiousatperiodarray(Iteration, testperiod)
        snbirds_alive = flocksizearray(Iteration, 1) - disease_Mort_atperiodarray(Iteration, testperiod)
        lprevaldbl = s_infectious / snbirds_alive
        detected = False
        pcrpositive = 0
        'normal mortality
        Snormal = Normal
        sdisease = disease

        If N_pcr_on_day = 2 Then
            If Normal + disease >= 2 * targetswabs Then
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
                templbirdstobetested = targetswabs - snumberdintube - n_norm_in_tube
                nposlivebirds_pertube = sim_Binomial(templbirdstobetested, lprevaldbl)
                If snumberdintube > 0 Or nposlivebirds_pertube > 0 Then
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
                templbirdstobetested = targetswabs - snumberdintube - n_norm_in_tube
                nposlivebirds_pertube = sim_Binomial(templbirdstobetested, lprevaldbl)
                If snumberdintube > 0 Or nposlivebirds_pertube > 0 Then
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
            templbirdstobetested = targetswabs - snumberdintube - n_norm_in_tube
            nposlivebirds_pertube = sim_Binomial(templbirdstobetested, lprevaldbl)
            If snumberdintube > 0 Or nposlivebirds_pertube > 0 Then
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
            If Normal + disease >= 3 * targetswabs Then
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
                templbirdstobetested = targetswabs - snumberdintube - n_norm_in_tube
                nposlivebirds_pertube = sim_Binomial(templbirdstobetested, lprevaldbl)
                If snumberdintube > 0 Or nposlivebirds_pertube > 0 Then
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
                templbirdstobetested = targetswabs - snumberdintube - n_norm_in_tube
                nposlivebirds_pertube = sim_Binomial(templbirdstobetested, lprevaldbl)
                If snumberdintube > 0 Or nposlivebirds_pertube > 0 Then
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
                templbirdstobetested = targetswabs - snumberdintube - n_norm_in_tube
                nposlivebirds_pertube = sim_Binomial(templbirdstobetested, lprevaldbl)
                If snumberdintube > 0 Or nposlivebirds_pertube > 0 Then
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

        If isonetubeper50 Then
            If N_pcr_on_day > 0 Then
                Dim N_extra_tubes As Long
                N_extra_tubes = Int((Normal + disease) / 50)
                For i = 1 To N_extra_tubes
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
                        Exit For
                    End If
                Next
            End If
        End If



        'End If
        If detected = True Then
            performonedaysurveillance_withlive = True
        End If
    End Function

    Private Function performonedaysurveillance(ByVal Normal As Double, ByVal disease As Integer, ByVal N_pcr_on_day As Long) As Boolean
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
            If Normal + disease >= 2 * targetswabs Then
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
            If Normal + disease >= 3 * targetswabs Then
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

        If isonetubeper50 Then
            If N_pcr_on_day > 0 Then
                Dim N_extra_tubes As Long
                N_extra_tubes = Int((Normal + disease) / 50)
                For i = 1 To N_extra_tubes
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
                        Exit For
                    End If
                Next
            End If
        End If



        'End If
        If detected = True Then
            performonedaysurveillance = True
        End If
    End Function
    Public Function hypergeometric(ByVal normal As Long, ByVal disease As Long, ByVal nswabs As Long)
        Try
            If normal + disease > 0 Then
                hypergeometric = myconnector.Evaluate("rhyper(1," & disease & "," & normal & "," & nswabs & ")")
            Else

                hypergeometric = 0

            End If
        Catch e As Exception
        End Try
    End Function
    Public Function sim_Binomial(ByVal N As Double, ByVal p As Double)
        Dim temp As Object
        temp = myconnector.Evaluate("rbinom(1," & N & "," & p & ")")
        sim_Binomial = temp
    End Function
    Public Sub main()
        form1.starttime = DateAndTime.Now
        noiterations = form1.niterations
        form1.calc_nperiod()
        nperiodsinaday = Math.Round(24 / form1.tau)
        initate_excel()
        initiatearrays()

        initiate_mortalitysim()
        Dim i As Long
        For i = 1 To iterend - iterstart + 1
            perform_one_flock_surveillance(i)
        Next
        writeresults()
        form1.endtime = DateAndTime.Now
        suvrunparamterrecorder()
        myworkbook.Save()
        If Not autoclose Then
            Dim msresult As Microsoft.VisualBasic.MsgBoxResult
            msresult = MsgBox("finshed close excel?", MsgBoxStyle.YesNoCancel)
            If msresult = 1 Then
                myexcel.Quit()
            Else
                myworkbook.Activate()
                myexcel.Visible = True
            End If
        Else
            myworkbook.Close()
            myexcel.Quit()
        End If



    End Sub
    Public Function Choose_diseaseday_on_movement_day() As Long
        'Disease_day_on_movement_day = myconnector.Evaluate("sample(" & movement_day_lower & ":" & movement_day_upper & ", size = 1)")
        Disease_day_on_movement_day = form1.Givendiseasedayonmovment

    End Function
    Public Sub writeresults()
        myworksheet = myworkbook.Worksheets("survpar")
        myworksheet.Range(myworksheet.Cells(3, 6), myworksheet.Cells(3 + iterend - iterstart + 1, 6)).Value = detectdayarray
        myworksheet.Range(myworksheet.Cells(3, 7), myworksheet.Cells(3 + iterend - iterstart + 1, 7)).Value = movementdayarray
        myworksheet.Range(myworksheet.Cells(3, 8), myworksheet.Cells(3 + iterend - iterstart + 1, 8)).Value = infectiousattesting
        myworksheet.Range(myworksheet.Cells(3, 9), myworksheet.Cells(3 + iterend - iterstart + 1, 9)).Value = infectiousatmovement
        myworksheet.Range(myworksheet.Cells(3, 10), myworksheet.Cells(3 + iterend - iterstart + 1, 10)).Value = egg_detectarray
        myworksheet.Range(myworksheet.Cells(3, 11), myworksheet.Cells(3 + iterend - iterstart + 1, 11)).Value = mort_detectarray
        'myworksheet.Range(myworksheet.Cells(3, 16), myworksheet.Cells(3 + noiterations, 16)).Value = egg_moved_array
        myworksheet.Range(myworksheet.Cells(3, 100), myworksheet.Cells(3 + iterend - iterstart + 1, 100)).Value = min_min_detectarray
        myworksheet.Cells(3, 53) = (iterend - iterstart + 1 - not_detect_count) / (iterend - iterstart + 1)
        myworksheet.Cells(3, 52) = (niterations_with_infectiousbirds - not_detect_count_withinfectiousbirds) / niterations_with_infectiousbirds
        myworksheet.Cells(3, 54) = getaverage(infe_innot_detct)
        myworksheet.Cells(3, 55) = getquantile(infe_innot_detct, 0.05)
        myworksheet.Cells(3, 56) = getquantile(infe_innot_detct, 0.95)
    End Sub
    Public Sub getcontam_egg(ByVal Curr_iter)
        'Dim k As Integer
        'Dim temeggrate As Double
        'temeggrate = myconnector.Evaluate("runif(1," & eggratelower & "," & eggratehigher & ")")
        'ReDim temp_diseaseeggproductarr(15)
        'For k = 1 To 14
        '    temp_diseaseeggproductarr(k) = contam_egg_array(Curr_iter, k) * temeggrate * (1 - dropinegg)
        'Next
    End Sub
    ''' <summary>
    ''' Model runs in terms of disease day. This method performs surveillance for one day. The actual movement day can be any disease day due to translation.
    ''' This translation factor is randomly selected with plausible values
    ''' </summary>
    ''' <param name="curr_iter"></param>
    ''' <remarks></remarks>
    Public Sub perform_one_flock_surveillance(ByVal curr_iter As Long)
        Dim normmortality As Double
        Dim mordetect As Boolean
        Dim actdayconeggs(15) As Single
        Dim testdayarray(15) As Integer
        Dim testacdayarray(15) As Integer
        Dim totaleggstobemoved As Single
        Dim diseasetimetilltotal As Single
        Dim acttodisease(15) As Integer
        Dim actdetectday As Integer
        Dim actdayindex As Integer
        Dim eggmovestartday As Integer
        Dim eggmoveendday As Integer
        Dim mycase As Integer
        mordetect = True
        Choose_diseaseday_on_movement_day()
        getnormalmortality(curr_iter, Disease_day_on_movement_day)

        get_temp_disease_mort(curr_iter, Disease_day_on_movement_day)
        getcontam_egg(curr_iter)
        Dim tempdifference As Integer


        mycase = 0

        For q = 1 To 15
            testdayarray(q) = 0
            testacdayarray(q) = 0
        Next q
        Dim minonestr, min2str, acstr, dailyteststr As String
        minonestr = ""
        min2str = ""
        acstr = ""
        dailyteststr = ""


        If mycase = 0 Then
            If is_daily_PCRtesting Then
                For q = 1 To Disease_day_on_movement_day
                    testdayarray(q) = 1
                    dailyteststr = "daily PCR testing"
                Next q

            Else
                dailyteststr = "no daily PCR testing"
            End If
            If is_daily_ACtesting Then
                For q = 1 To Disease_day_on_movement_day
                    testacdayarray(q) = 1
                    dailyteststr = dailyteststr & "AC testing"
                Next q
            Else
                dailyteststr = dailyteststr & "no daily AC testing "
            End If



            If Not is_daily_PCRtesting Then
                Dim movedaystr As String
                movedaystr = 2 'change here for movement day
                testdayarray(Disease_day_on_movement_day) = movedaystr
            End If

            ''''Antigen capture''''
            If Not is_daily_ACtesting Then
                acstr = 0 'set antigen capture test
                testacdayarray(Disease_day_on_movement_day) = acstr
            End If

            '''''''''''''''
            If Not is_daily_PCRtesting Then
                minonestr = 0 ''Change here for day before movement
                If Disease_day_on_movement_day - 1 >= 1 Then
                    testdayarray(Disease_day_on_movement_day - 1) = minonestr
                End If
            End If
            If Not is_daily_PCRtesting Then
                min2str = 0
                If Disease_day_on_movement_day - 2 >= 1 Then
                    testdayarray(Disease_day_on_movement_day - 2) = min2str
                End If
            End If

            'If randommovementday + 2 <= 14 Then
            '    testdayarray(randommovementday + 2) = 1
            'End If
            'If randommovementday + 1 <= 14 Then
            '    testdayarray(randommovementday + 1) = 1
            'End If
            eggmovestartday = Disease_day_on_movement_day - ndaysmovedpershipment - minhldtime + 1
            If eggmovestartday < 1 Then eggmovestartday = 1
            eggmoveendday = Disease_day_on_movement_day - minhldtime
            If eggmoveendday < 1 Then eggmoveendday = 1

            Dim flockstring As String
            If form1.is_turkey_hen Then
                flockstring = "turkey hen "
            Else
                flockstring = "turkey tom "

            End If
            If Not is_daily_PCRtesting Then
                specialcommentstring = flockstring & testdayarray(Disease_day_on_movement_day).ToString & " pooled tests on movement day and " & minonestr & " pooled tests on day before and upto " & targetswabs & " swabs per pool and " & min2str & "  tests two days before"
            Else
                specialcommentstring = flockstring & " with " & dailyteststr
            End If

            If Not is_daily_PCRtesting Then
                specialcommentstring_2 = acstr.ToString & "  flockside antigen capture tests and with " & ACtargetswabs & " swabs with actests pooled being" & is_Ac_test_pooled
            Else
                specialcommentstring_2 = "Broilers with" & dailyteststr
            End If

        Else

        End If

        Dim detected As Boolean
        Dim eggdetected As Boolean
        Dim mortdetected As Boolean
        Dim firstdetectflag As Boolean
        Dim firsteggdetectflag As Boolean
        Dim firstmortdetectflag As Boolean

        detected = False
        eggdetected = False
        firstdetectflag = False
        firsteggdetectflag = False
        firstmortdetectflag = False
        Dim disease_detectday As Integer
        Dim egg_detectday As Integer
        Dim mort_detectday As Integer
        disease_detectday = 0

        Dim i As Integer
        For i = 1 To 14

            Dim testpcr, testac As Integer

            testpcr = testdayarray(i)
            testac = testacdayarray(i)
            'testpcr = False


            'PCR detection
            If detected = False Then
                ' detected = performonedaysurveillance(tempdailymortalityarr(i - 1), tempdiseasemortalityarr(i), testpcr)
                detected = performonedaysurveillance_withlive(tempdailymortalityarr(i - 1), tempdiseasemortalityarr(i), testpcr, curr_iter, (i * nperiodsinaday + 1))
            End If

            'Code below for ac testing about midnight or 18 hours after sampling of normal mortality
            Dim temp_tempd As Double
            temp_tempd = myconnector.Evaluate("runif(1," & 0.6666 & "," & 0.66666666 & ")")
            '
            Dim tempacmort, temdailyacmort, t_moveinteger As Long
            'Creating a new pool for the ac testimg mortality
            t_moveinteger = Math.Round((temp_tempd * 24 / form1.tau))
            get_temp_ac_sp_mortality(curr_iter, i, t_moveinteger)
            temdailyacmort = tempdailymortalityarr(i) * (t_moveinteger * form1.tau / 24)
            If detected = False Then
                detected = performoneACTomdaysurveillance(temdailyacmort, temp_acmortality, testac)
            End If
            'Drop in egg detection'
            'If eggdetected = False Then
            '    eggdetected = perform_egg_surv(tempeggratearr(i))
            'End If
            'mortality detection
            If mortdetected = False Then
                mortdetected = perform_mort_surv(tempdailymortalityarr(i - 1), tempdiseasemortalityarr(i), flocksizearray(curr_iter, 1))
            End If

            If detected Then
                If firstdetectflag = False Then
                    disease_detectday = i
                    firstdetectflag = True
                    detected = True
                End If
            End If
            'If eggdetected Then
            '    If firsteggdetectflag = False Then
            '        firsteggdetectflag = True
            '        eggdetected = True
            '        egg_detectday = i
            '    End If
            'End If
            If mortdetected Then
                If firstmortdetectflag = False Then
                    firstmortdetectflag = True
                    mortdetected = True
                    mort_detectday = i
                End If
            End If

        Next

        If detected = False Then
            disease_detectday = 15

        End If
        If eggdetected = False Then
            egg_detectday = 15

        End If
        If mortdetected = False Then
            mort_detectday = 15
        End If



        'For i = eggmovestartday To eggmoveendday
        '    totaleggstobemoved = totaleggstobemoved + temp_diseaseeggproductarr((i))
        'Next


        Dim tempmindetect_time As Integer
        If disease_detectday >= mort_detectday Then tempmindetect_time = mort_detectday Else tempmindetect_time = disease_detectday





        detectdayarray(curr_iter - 1, 0) = disease_detectday
        movementdayarray(curr_iter - 1, 0) = Disease_day_on_movement_day


        If tempmindetect_time > Disease_day_on_movement_day Then
            infectiousattesting(curr_iter - 1, 0) = (infectiousatperiodarray(curr_iter, Math.Round(Math.Max(Disease_day_on_movement_day - 2, 0) * (24 / form1.tau) + 1))) / flocksizearray(curr_iter, 1)
        Else
            infectiousattesting(curr_iter - 1, 0) = 0
        End If
        If tempmindetect_time > Disease_day_on_movement_day Then
            infectiousatmovement(curr_iter - 1, 0) = (infectiousatperiodarray(curr_iter, Math.Round(Disease_day_on_movement_day * (24 / form1.tau) + 3))) / flocksizearray(curr_iter, 1)
        Else
            infectiousatmovement(curr_iter - 1, 0) = 0
        End If


        'special analysis

        If (infectiousatperiodarray(curr_iter, Math.Round(Disease_day_on_movement_day * (24 / form1.tau) + 3)) > 0 Or infectiousatperiodarray(curr_iter, Math.Round(13 * (24 / form1.tau) + 3)) > 0) Then
            niterations_with_infectiousbirds += 1
            If tempmindetect_time > Disease_day_on_movement_day Then
                not_detect_count_withinfectiousbirds += 1
            End If
        End If

        If tempmindetect_time > Disease_day_on_movement_day Then
            not_detect_count = not_detect_count + 1
            ReDim Preserve infe_innot_detct(0, not_detect_count - 1)
            infe_innot_detct(0, not_detect_count - 1) = (infectiousatperiodarray(curr_iter, Math.Round(Disease_day_on_movement_day * (24 / form1.tau) + 3))) / flocksizearray(curr_iter, 1)
        End If


        'egg_detectarray(curr_iter - 1, 0) = egg_detectday
        mort_detectarray(curr_iter - 1, 0) = mort_detectday
        'If tempmindetect_time > randommovementday Then
        '    egg_moved_array(curr_iter - 1, 0) = totaleggstobemoved
        'Else
        '    egg_moved_array(curr_iter - 1, 0) = 0
        'End If
        min_min_detectarray(curr_iter - 1, 0) = tempmindetect_time
        If curr_iter = 999 Then

            Beep()
        End If

    End Sub




End Class
