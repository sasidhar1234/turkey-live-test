''' <summary>
''' This model can take n hns infected via semen or a single infected hen.
''' if infection is via semen, then the dynamics of that spread need to be entered in mannually in AI putput file based on run of tom surveillance.
''' 27 of survpar sheet of AI output should have info on number of infectious toms at movement
''' 33 and 34 should have info on detected by movement day+2 and 3
''' if a tom flock is infected then movement of hatching eggs from associated henflocks are restricted.
''' </summary>
''' <remarks></remarks>
Public Class hensurvsim
    Public noiterations As Long
    Public outputstringname As String
    Dim myexcel As Microsoft.Office.Interop.Excel.Application
    Dim myworkbook As Microsoft.Office.Interop.Excel.Workbook
    Dim myworksheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim myconnector As STATCONNECTORSRVLib.StatConnector
    Dim mymortsimulator As Mortalitydatasimulator
    Dim diseasemortalityarray(,) As Object
    Dim infectiousatperiodarray(,) As Object
    Dim egg_productrate_period_array(,) As Object
    Dim contam_egg_array(,) As Object
    Dim eggratelower As Single = 0.45
    Dim eggratehigher As Single = 0.7
    Dim ndaystorun As Long
    Dim tempdailymortalityarr() As Double
    Dim tempdiseasemortalityarr() As Double
    Dim tempeggratearr() As Double
    Dim flocksizearray(,) As Object
    Dim eggstartratearray(,) As Object
    Dim tomdetectdayplustwo(,) As Object
    Dim tomdetectdayplus_three(,) As Object
    Dim nlatentfor_external_eggs(,) As Object
    Dim temp_diseaseeggproductarr() As Double
    Dim mortfraclimt As Double = 0.002
    Dim eggdecreaslimittwice As Double = 0.15
    Dim eggdecreaselimitonce As Double = 0.15
    Dim targetswabs As Long = 5
    Dim sensitvity As Double = 0.865
    Dim acsensitvity As Double = 0.7
    Dim dropinegg As Double = 0.75
    Dim pcrcounter As Long
    Dim detectdayarray(,) As Double
    Dim egg_detectarray(,) As Double
    Dim mort_detectarray(,) As Double
    Dim tom_detectarray(,) As Double
    Dim min_min_detectarray(,) As Double
    Dim movementdayarray(,) As Double
    Dim infectiousattesting(,) As Double
    Dim infectiousatmovement(,) As Double
    Dim detectedbeforemovement As Boolean
    Dim egg_moved_array(,) As Double
    Dim external_egg_movedarray(,) As Double
    Dim minhldtime As Integer = 2
    Dim movement_day_lower As Long = 2
    Dim movement_day_upper As Long = 9
    Dim Disease_day_when_movement_occurs As Long
    Dim ndaysmovedpershipment As Integer = 5
    Dim specialcommentstring As String
    Public isonetubeper50 As Boolean = True

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
    Public Sub suvrunparamterrecorder()
        myworksheet = myworkbook.Worksheets("Runp")
       
        myworksheet.Cells(18, 2) = Form1.niterations
       
        myworksheet.Cells(21, 2) = getaverage(flocksizearray)
        myworksheet.Cells(22, 2) = Form1.starttime
        myworksheet.Cells(23, 2) = Form1.endtime
        myworksheet.Cells(24, 2) = DateAndTime.DateDiff(DateInterval.Second, Form1.starttime, Form1.endtime)
        myworksheet.Cells(25, 2) = minhldtime
        myworksheet.Cells(26, 2) = movement_day_upper
        myworksheet.Cells(27, 2) = movement_day_lower
        myworksheet.Cells(28, 3) = specialcommentstring
    End Sub

    Public Sub New(ByVal con As STATCONNECTORSRVLib.StatConnector)
        myconnector = con
        noiterations = Form1.niterations
    End Sub
    Public Sub initate_excel()
        Try
            myexcel = New Microsoft.Office.Interop.Excel.Application
            myworkbook = myexcel.Workbooks.Open("C:\Users\public\Documents\Work\Turkey transmission\breeder transmission\Modeling of Pensylvania\pensylvania\AI Output.xlsx")
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
        myworksheet = myworkbook.Worksheets("Egg")
        eggstartratearray = myworksheet.Range(myworksheet.Cells(3, 80), myworksheet.Cells(3 + noiterations - 1, 80)).Value
        myworksheet = myworkbook.Worksheets("Survpar")
        tomdetectdayplustwo = myworksheet.Range(myworksheet.Cells(3, 33), myworksheet.Cells(3 + noiterations - 1, 33)).Value
        tomdetectdayplus_three = myworksheet.Range(myworksheet.Cells(3, 34), myworksheet.Cells(3 + noiterations - 1, 34)).Value
        myworksheet = myworkbook.Worksheets("Infectious")
        infectiousatperiodarray = myworksheet.Range(myworksheet.Cells(2, 2), myworksheet.Cells(2 + noiterations - 1, 61)).Value
        myworksheet = myworkbook.Worksheets("clinical infec")
        contam_egg_array = myworksheet.Range(myworksheet.Cells(2, 65), myworksheet.Cells(2 + noiterations - 1, 78)).Value
        myworksheet = myworkbook.Worksheets("Egg")
        egg_productrate_period_array = myworksheet.Range(myworksheet.Cells(2, 2), myworksheet.Cells(2 + noiterations - 1, 61)).Value
        myworksheet = myworkbook.Worksheets("Latent")
        nlatentfor_external_eggs = myworksheet.Range(myworksheet.Cells(2, 2), myworksheet.Cells(2 + noiterations - 1, 2)).Value



        ReDim detectdayarray(noiterations, 0)
        ReDim movementdayarray(noiterations, 0)
        ReDim infectiousatmovement(noiterations, 0)
        ReDim infectiousattesting(noiterations, 0)
        ReDim egg_moved_array(noiterations, 0)
        ReDim egg_detectarray(noiterations, 0)
        ReDim mort_detectarray(noiterations, 0)
        ReDim tom_detectarray(noiterations, 0)
        ReDim external_egg_movedarray(noiterations, 0)
        ReDim min_min_detectarray(noiterations, 0)


    End Sub
    Public Sub get_temp_disease_mort(ByVal curr_iter As Long)
        Dim k As Integer
        ReDim tempdiseasemortalityarr(15)
        For k = 1 To 14
            tempdiseasemortalityarr(k) = diseasemortalityarray(curr_iter, k)
        Next
    End Sub
    Public Sub get_temp_eggrate(ByVal curr_iter As Long)
        Dim k, p As Integer
        Dim tempsum As Double
        ReDim tempeggratearr(15)
        For k = 1 To 14
            tempsum = 0
            For p = 1 To 4
                tempsum = tempsum + egg_productrate_period_array(curr_iter, k * Math.Round(24 / Form1.tau) - (p - 1))
            Next
            tempeggratearr(k) = tempsum / 4
        Next
    End Sub

    Private Function perform_egg_surv(ByVal eggproductrate1, ByVal eggproductionrate2, ByVal eggproductionrate3) As Boolean
        Dim flag As Boolean = False
        '' modifying here for .15

        If eggproductrate1 - eggproductionrate3 >= eggdecreaslimittwice Then
            flag = True
        End If


        If eggproductionrate2 - eggproductionrate3 >= eggdecreaselimitonce Then
            flag = True
        End If

        perform_egg_surv = flag


    End Function

    Private Function perform_mort_surv(ByVal Normal As Double, ByVal disease As Integer, ByVal flockssize As Long) As Boolean
        If Normal + disease > flockssize * mortfraclimt Then
            perform_mort_surv = True
        Else
            perform_mort_surv = False
        End If
    End Function

    Private Function performoneACTomdaysurveillance(ByVal Normal As Double, ByVal disease As Integer, ByVal numberoftubes As Integer, ByVal mordetect As Boolean, ByVal mordetectdis As Integer, ByVal everybody As Boolean, ByVal doactest As Long, ByVal flockssize As Long) As Boolean
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

        If doactest = 1 Then
            detected = False
            pcrpositive = 0
            'normal mortality
            Snormal = Normal
            sdisease = disease

            If disease >= 5 Then
                myfirstn = 5
            Else
                myfirstn = disease
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


    End Function


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
        Form1.starttime = DateAndTime.Now


        noiterations = Form1.niterations
        initate_excel()
        initiatearrays()
        initiatemortalitysim()
        Dim i As Long
        For i = 1 To noiterations
            perform_one_flock_surveillance(i)
        Next

        writeresults()
        Form1.endtime = DateAndTime.Now
        suvrunparamterrecorder()
        myworkbook.Save()
        myexcel.Quit()
        MsgBox("finished")





    End Sub
    Public Function createrandommovementday() As Long
        If Form1.infectedbysemen Then
            Disease_day_when_movement_occurs = 3
        Else
            Disease_day_when_movement_occurs = myconnector.Evaluate("sample(" & movement_day_lower & ":" & movement_day_upper & ", size = 1)")
        End If

        'randommovementday = 1
    End Function
    Public Sub writeresults()
        myworksheet = myworkbook.Worksheets("survpar")
        myworksheet.Range(myworksheet.Cells(3, 6), myworksheet.Cells(3 + noiterations, 6)).Value = detectdayarray
        myworksheet.Range(myworksheet.Cells(3, 7), myworksheet.Cells(3 + noiterations, 7)).Value = movementdayarray
        myworksheet.Range(myworksheet.Cells(3, 8), myworksheet.Cells(3 + noiterations, 8)).Value = infectiousattesting
        myworksheet.Range(myworksheet.Cells(3, 9), myworksheet.Cells(3 + noiterations, 9)).Value = infectiousatmovement
        myworksheet.Range(myworksheet.Cells(3, 10), myworksheet.Cells(3 + noiterations, 10)).Value = egg_detectarray
        myworksheet.Range(myworksheet.Cells(3, 11), myworksheet.Cells(3 + noiterations, 11)).Value = mort_detectarray
        myworksheet.Range(myworksheet.Cells(3, 16), myworksheet.Cells(3 + noiterations, 16)).Value = egg_moved_array
        myworksheet.Range(myworksheet.Cells(3, 19), myworksheet.Cells(3 + noiterations, 19)).Value = tom_detectarray
     
        myworksheet.Range(myworksheet.Cells(3, 100), myworksheet.Cells(3 + noiterations, 100)).Value = min_min_detectarray
    End Sub
    Public Sub getcontam_egg(ByVal Curr_iter)
        Dim k As Integer
        Dim temeggrate As Double
        'temeggrate = myconnector.Evaluate("runif(1," & eggratelower & "," & eggratehigher & ")")
        ReDim temp_diseaseeggproductarr(15)
        For k = 1 To 14
            temp_diseaseeggproductarr(k) = contam_egg_array(Curr_iter, k) * tempeggratearr(k) * (1 - dropinegg)
        Next
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
        getnormalmortality(curr_iter)
        createrandommovementday()
        get_temp_eggrate(curr_iter)
        get_temp_disease_mort(curr_iter)
        getcontam_egg(curr_iter)

        Dim tempdifference As Integer


        mycase = 0

        For q = 1 To 15
            testdayarray(q) = 0
            testacdayarray(q) = 0
        Next q
        specialcommentstring = "hen infected from toms baseline so testing pcr two days before movement vbcrlf specifally movement day -1 and -2 testday arrays are set to 1"
        If mycase = 0 Then
            testdayarray(Disease_day_when_movement_occurs) = 1

            ''''Antigen capture''''
            ' testacdayarray(randommovementday) = 1
            '''''''''''''''

            If Disease_day_when_movement_occurs - 1 >= 1 Then
                testdayarray(Disease_day_when_movement_occurs - 1) = 1
            End If
            'If randommovementday - 2 >= 1 Then
            '    testdayarray(randommovementday - 2) = 1
            'End If
            'If randommovementday + 2 <= 14 Then
            '    testdayarray(randommovementday + 2) = 1
            'End If
            'If randommovementday + 1 <= 14 Then
            '    testdayarray(randommovementday + 1) = 1
            'End If
            eggmovestartday = Disease_day_when_movement_occurs - ndaysmovedpershipment - minhldtime + 1
            If eggmovestartday < 1 Then eggmovestartday = 1
            eggmoveendday = Disease_day_when_movement_occurs - minhldtime
            If eggmoveendday < 1 Then eggmoveendday = 1
        Else

        End If


        '
       




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
        Dim tomdetectday As Integer


        disease_detectday = 0

        Dim i As Integer
        For i = 1 To 14

            Dim testpcr, testac As Integer

            testpcr = testdayarray(i)
            testac = testacdayarray(i)
            'testpcr = False

            Dim mordetdismor As Integer
            mordetdismor = tempdailymortalityarr(i)
            'PCR detection
            If detected = False Then
                detected = performonedaysurveillance(tempdailymortalityarr(i - 1), tempdiseasemortalityarr(i), 2, True, mordetdismor, False, testpcr, Form1.nbirds)
            End If

            If detected = False Then
                detected = performoneACTomdaysurveillance(tempdailymortalityarr(i - 1), tempdiseasemortalityarr(i), 2, True, mordetdismor, False, testac, Form1.nbirds)
            End If
            'Drop in egg detection'
            If eggdetected = False Then
                Dim eggrate1, eggrate2, eggrate3 As Double
                eggrate3 = tempeggratearr(i)
                If i >= 2 Then eggrate2 = tempeggratearr(i - 1) Else eggrate2 = eggstartratearray(curr_iter, 1)
                If i >= 3 Then eggrate1 = tempeggratearr(i - 2) Else eggrate1 = eggstartratearray(curr_iter, 1)
                eggdetected = perform_egg_surv(eggrate1, eggrate2, eggrate3)
            End If
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
            If eggdetected Then
                If firsteggdetectflag = False Then
                    firsteggdetectflag = True
                    eggdetected = True
                    egg_detectday = i
                End If
            End If
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
        tomdetectday = 15

        If Form1.infectedbysemen = True Then
            If tomdetectdayplustwo(curr_iter, 1) < 1 Then
                tomdetectday = 2
            ElseIf tomdetectdayplus_three(curr_iter, 1) < 1 Then
                tomdetectday = 3
            End If
        End If
        For i = eggmovestartday To eggmoveendday
            totaleggstobemoved = totaleggstobemoved + temp_diseaseeggproductarr((i))
        Next


        Dim tempmindetect_time As Integer
        If disease_detectday >= mort_detectday Then
            tempmindetect_time = mort_detectday
        Else
            tempmindetect_time = disease_detectday
        End If
        If egg_detectday < tempmindetect_time Then
            tempmindetect_time = egg_detectday
        End If
        If Form1.infectedbysemen Then
            If tomdetectday < tempmindetect_time Then
                tempmindetect_time = tomdetectday
            End If
        End If


        detectdayarray(curr_iter - 1, 0) = disease_detectday
        movementdayarray(curr_iter - 1, 0) = Disease_day_when_movement_occurs
        tom_detectarray(curr_iter - 1, 0) = tomdetectday
        min_min_detectarray(curr_iter - 1, 0) = tempmindetect_time
        If tempmindetect_time > Disease_day_when_movement_occurs Then
            infectiousattesting(curr_iter - 1, 0) = infectiousatperiodarray(curr_iter, Math.Round(Math.Max(Disease_day_when_movement_occurs - 2, 0) * (24 / Form1.tau) + 1))
        Else
            infectiousattesting(curr_iter - 1, 0) = 0
        End If
        If tempmindetect_time > Disease_day_when_movement_occurs Then
            infectiousatmovement(curr_iter - 1, 0) = infectiousatperiodarray(curr_iter, Math.Round(Disease_day_when_movement_occurs * (24 / Form1.tau) + 1))
        Else
            infectiousatmovement(curr_iter - 1, 0) = 0
        End If
        egg_detectarray(curr_iter - 1, 0) = egg_detectday
        mort_detectarray(curr_iter - 1, 0) = mort_detectday
        If Form1.infectedbysemen Then
            If tempmindetect_time > Disease_day_when_movement_occurs Then
                egg_moved_array(curr_iter - 1, 0) = nlatentfor_external_eggs(curr_iter, 1) * tempeggratearr(1) * (1 - dropinegg) / 2
            Else
                egg_moved_array(curr_iter - 1, 0) = 0
            End If

        Else
            If tempmindetect_time > Disease_day_when_movement_occurs Then
                egg_moved_array(curr_iter - 1, 0) = totaleggstobemoved
            Else
                egg_moved_array(curr_iter - 1, 0) = 0
            End If

        End If


        If curr_iter = 999 Then
            Beep()
        End If

    End Sub




End Class
