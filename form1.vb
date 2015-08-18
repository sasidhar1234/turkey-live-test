Public Class form1
    ''start copy
    Dim myspeech As Speech.Synthesis.SpeechSynthesizer
    Public nbirds As Integer = 4000
    Public niterations As Integer = 6000
    Public updateform As Integer = 100
    Public npens As Integer = 1
    Public contactrate As Single = 0.5 'contact rate per period
    Public contactrateotherpen As Single = 3
    Public tau As Single = 2.4
    Public latentdistralpha As Double
    Public eggstarttime As Double = 24
    Public latentdistrbeta As Double
    Public ndaystorun As Integer = 15
    Public subclin_dist_alpha As Double
    Public subclin_dist_beta As Double
    Public clin_dist_alpha As Double
    Public clin_dist_beta As Double
    Public current_time As Double = 1
    Public current_iter As Long = 1
    Dim pencapacity As Integer = 290000
    Dim pens() As barn
    Dim cowlatenttimearray() As Double
    Dim cowsublintimearray() As Double
    Dim cowinftimearray() As Double
    Dim cowclintimearray() As Double
    Dim flocksizearray() As Double
    Dim nsemeninfecarray(,) As Object
    Dim myconnector As New STATCONNECTORSRVLib.StatConnector
    Dim pdevice1 As New STATCONNECTORCLNTLib.StatConnectorGraphicsDevice
    Public nperiods As Long
    Public output_nsusceptible()
    Public output_nlatent()
    Public output_ninfectious()
    Public output_nremoved()
    Public output_nsubclincal()
    Public output_nclincalinfectious()
    Public Output_nclinical_removed()
    Public Output_egg_rate_inperiod()
    Dim myexcel As Microsoft.Office.Interop.Excel.Application
    Dim myworkbook As Microsoft.Office.Interop.Excel.Workbook
    Dim myworksheet As Microsoft.Office.Interop.Excel.Worksheet
    Public infectedbysemen As Boolean = False
    Public weeklytwicetom As Boolean = False
    Public starttime As Date
    Public endtime As Date
    Dim eggratelower As Single = 0.45
    Dim eggratehigher As Single = 0.7
    Dim starting_egg_rate_array() As Double
    Dim temp_egg_rate_iter As Double
    Dim dropinegg As Double = 0.75
    Public iscontactratestochastic As Boolean = True
    Public contactratearray() As Double
    Public stochcontactratehigher As Single = 12
    Public stochasticcontactratelower As Single = 1
    Public tempcontactrateiter As Double
    Public Givendiseasedayonmovment As Long = 4
    Public survsimparametertochange As Long
    Public is_turkey_hen As Boolean = False
    Public excelwritingtime As TimeSpan
    Public barniteratingtime As TimeSpan
    Public distributionloadingtime As TimeSpan
    Public arraycreatingtime As TimeSpan
    Public Sub initate_excel()
        Try
            myexcel = New Microsoft.Office.Interop.Excel.Application
            myworkbook = myexcel.Workbooks.Open("C:\Codes\Live turkey Movement\AIoutput\AI Output.xlsx")
        Catch ex As Exception
            myexcel.Quit()
            MsgBox(ex)
        End Try
    End Sub
    Public Sub close_excel()
        myworkbook.Close(True)
        myexcel.Quit()
    End Sub
    Public Sub loadstochasticarrays()
        Try
            If iscontactratestochastic Then
                'contactratearray = myconnector.Evaluate("runif(" & niterations & "," & stochasticcontactratelower & "," & stochcontactratehigher & ")")
                'Brians noaprioribeta, person V or inverse gamma distribution
                'includes adjustment for time steps
                ''______________________________________________________________________________________________________________________________
                contactratearray = myconnector.Evaluate("(1/(rgamma(" & niterations & ",shape=16.775,scale =1/68.581)))/" & Math.Round(24 / Me.tau))
                ''__________________________________________________________________________________________________________________________________
                'truncation
                'senstivity analysis contact rate
                '-_______________________________________________________________________________________________________________________________________________
                contactratearray = myconnector.Evaluate("runif(" & niterations & "," & stochasticcontactratelower & "," & stochcontactratehigher & ")/" & Math.Round(24 / Me.tau))
                '_
                Dim k, z, ncontacts As Integer
                Dim dumcarray(3) As Integer
                dumcarray(0) = 2
                dumcarray(1) = 4
                dumcarray(2) = 6
                dumcarray(3) = 8
                ncontacts = 4
                Dim countperrate As Integer
                countperrate = CInt(niterations / (4))

                For z = 1 To ncontacts - 1
                    For k = 0 To countperrate - 1
                        contactratearray(countperrate * (z - 1) + k) = dumcarray(z - 1) / Math.Round(24 / Me.tau)
                    Next k
                Next
                For k = (ncontacts - 1) * countperrate To contactratearray.GetUpperBound(0)
                    contactratearray(k) = dumcarray(ncontacts - 1) / Math.Round(24 / Me.tau)
                Next k
                ' _______________________________________________________________________________________________________________________________________________________()


                For k = contactratearray.GetLowerBound(0) To contactratearray.GetUpperBound(0)
                    If contactratearray(k) < 0.025 Then contactratearray(k) = 0.025
                    If contactratearray(k) > 100 Then contactratearray(k) = 100
                Next
            Else
                'includes adjustment for timestep
                contactratearray = myconnector.Evaluate("runif(" & niterations & "," & contactrate & "," & contactrate & ")/" & Math.Round(24 / Me.tau))
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub initiate_statconnector()
        Try
            myconnector.Init("R")
            myconnector.AddGraphicsDevice("mygraphic", pdevice1)
        Catch l As Exception
        End Try
    End Sub
    Public Sub set_temp_egg_rate(ByVal current_iter)
        Dim k As Integer
        temp_egg_rate_iter = myconnector.Evaluate("runif(1," & eggratelower & "," & eggratehigher & ")")
        starting_egg_rate_array(current_iter - 1) = temp_egg_rate_iter
    End Sub

    Public Function calculate_egg_production_rate(ByVal n_infectious As Double, ByVal nremoved As Double) As Double
        If nbirds > nremoved Then
            calculate_egg_production_rate = ((n_infectious * temp_egg_rate_iter * (1 - dropinegg)) + ((nbirds - n_infectious - nremoved) * temp_egg_rate_iter)) / (nbirds - nremoved)
        Else
            calculate_egg_production_rate = 0
        End If

    End Function

    Public Sub loadflocksize()
        Try
            If is_turkey_hen Then
                'changed on sep 20th
                flocksizearray = myconnector.Evaluate("rlnorm(" & niterations & ", meanlog=9.4592, sdlog= 0.3445)")
                Dim k As Integer
                For k = flocksizearray.GetLowerBound(0) To flocksizearray.GetUpperBound(0)
                    If flocksizearray(k) < 1000 Then flocksizearray(k) = 1000
                    If flocksizearray(k) > 50000 Then flocksizearray(k) = 50000
                Next
            Else
                flocksizearray = myconnector.Evaluate("rexp(" & niterations & ",1/15188)")

                'flocksizearray = myconnector.Evaluate("rgamma(" & niterations & ",shape=7.88,scale =426)")
                Dim k As Integer
                For k = flocksizearray.GetLowerBound(0) To flocksizearray.GetUpperBound(0)
                    If flocksizearray(k) < 2000 Then flocksizearray(k) = 2000
                    If flocksizearray(k) > 40000 Then flocksizearray(k) = 40000
                Next
            End If


            ReDim starting_egg_rate_array(niterations)

        Catch l As Exception


        End Try

    End Sub


    Public Sub loadtimedistributionsr()
        Try
            'cowlatenttimearray = myconnector.Evaluate("rweibull(" & nbirds & ", shape = 1.8727,scale = 0.79096)*24")
            'cowlatenttimearray = myconnector.Evaluate("rnorm(" & nbirds & ",mean=0.81034,sd = 0.273)*24")
            '_______________________________________________________________________________________
            'Aldous Alone for HPAI H5N1 turkeys
            cowlatenttimearray = myconnector.Evaluate("rgamma(" & nbirds & ", shape =10.0316549,scale=0.1263083)*24")
            '__________________________________________________________________________________________
            '_______________________________________________________________________________________
            'Vandergoot and Saenz HPAI H5N2 turkeys
            'cowlatenttimearray = myconnector.Evaluate("rgamma(" & nbirds & ", shape = 8.05423260753173,scale=0.05)*24")
            '__________________________________________________________________________________________




            Dim k As Integer
            For k = cowlatenttimearray.GetLowerBound(0) To cowlatenttimearray.GetUpperBound(0)
                If cowlatenttimearray(k) < 0 Then cowlatenttimearray(k) = 0
                If cowlatenttimearray(k) > 72 Then cowlatenttimearray(k) = 72
            Next
            'modified longer
            'shorter 
            'cowinftimearray = myconnector.Evaluate("rlnorm(" & nbirds & ", meanlog = .579, sdlog= .568)*24")
            'for 48 hours
            'cowinftimearray = myconnector.Evaluate("(rlnorm(" & nbirds & ", meanlog = .196, sdlog = .629)+rnorm(" & nbirds & ", 2,.5))*24")
            'longer
            'cowinftimearray = myconnector.Evaluate("rgamma(" & nbirds & ", shape =8.3284,scale =0.51863)*24")
            '_______________________________________________________________________________________
            'Aldous Alone for HPAI H5N1 turkeys
            'cowinftimearray = myconnector.Evaluate("rweibull(" & nbirds & ", shape =1.1031827,scale=1.3289883)*24")
            '____________________________________________________________________________________
            'Pensylvania vandergoot and swayne and eggert based on chicken and conservatively considering a latent period between
            '0 and .25 days for the swayne and eggert data
            ' cowinftimearray = myconnector.Evaluate("rweibull(" & nbirds & ", shape =1.965111,scale=4.237553)*24")
            '__________________________________________________________________________________________
            ''Fujian H5N2
            cowinftimearray = myconnector.Evaluate("rweibull(" & nbirds & ", shape =4.203382,scale=3.803676)*24")

            For k = cowinftimearray.GetLowerBound(0) To cowinftimearray.GetUpperBound(0)
                If cowinftimearray(k) > 240 Then cowinftimearray(k) = 240
                ' If cowinftimearray(k) + cowlatenttimearray(k) < 72 Then cowinftimearray(k) = 72 - cowlatenttimearray(k)
            Next

            'cowsublintimearray = myconnector.Evaluate("rgamma(" & nbirds & ",shape=20,scale =0.05)*24")
            'cowclintimearray = myconnector.Evaluate("rgamma(" & nbirds & ",shape=100,scale =1)*24")

        Catch l As Exception


        End Try

    End Sub




    ''' <summary>
    ''' creates and initiallizes arrays that refer to prens and cows
    ''' 
    ''' </summary>
    ''' <param name="current_iter"></param>
    ''' <remarks></remarks>
    Public Sub createpensandcows(ByVal current_iter)
        Dim ncowsperpen As Integer
        Dim dummycow As bird
        Dim dummypen As barn

        ReDim pens(npens + 1)
        For j = 1 To npens
            dummypen = New barn
            dummypen.penid = j
            dummypen.capacity = pencapacity
            dummypen.susceptible = New ArrayList
            dummypen.latent = New ArrayList
            dummypen.infectious = New ArrayList
            dummypen.removed = New ArrayList
            dummypen.new_susceptible = New ArrayList
            dummypen.new_latent = New ArrayList
            dummypen.new_infectious = New ArrayList
            dummypen.new_removed = New ArrayList
            pens(j) = dummypen
        Next
        Dim pencowcounter = 0
        Dim currentpen As Integer
        currentpen = 1

        For i = 1 To nbirds
            dummycow = New bird(currentpen, i)
            dummycow.Lengthlat = cowlatenttimearray(i - 1)
            ' dummycow.lengthsubclin = cowsublintimearray(i - 1)
            dummycow.lengthinfectious = cowinftimearray(i - 1)
            ' dummycow.lengthclinical = cowclintimearray(i - 1)
            pencowcounter += 1
            pens(currentpen).susceptible.Add(dummycow)
            pens(currentpen).nsusceptible += 1
            If pencowcounter + 1 > pens(currentpen).capacity Then
                currentpen += 1
                pencowcounter = 0
            End If
        Next i
    End Sub

    Public Sub infectacow(ByVal current_iter)
        Dim i As Integer

        Dim ncowsinfected As Long
        ncowsinfected = 1

        If ncowsinfected > nbirds - 1 Then
            ncowsinfected = nbirds - 1

        End If
        '0.15 * Rnd() * n cows

        For i = 0 To ncowsinfected - 1
            Dim tempcow As bird
            tempcow = pens(1).susceptible(i)
            tempcow.timeofinfection = current_time
            pens(1).latent.Add(tempcow)
        Next
        pens(1).susceptible.RemoveRange(0, ncowsinfected)
        pens(1).nsusceptible = pens(1).nsusceptible - ncowsinfected
        pens(1).nlatent = pens(1).nlatent + ncowsinfected
    End Sub

    Public Function calc_ninfected_inperiod(ByVal S As Long, ByVal I As Long, ByVal N As Long, ByVal R As Long) As Integer
        Dim p As Double
        p = 1 - Math.Exp(CDbl(-(tempcontactrateiter * I) / (N - R - 1)))
        calc_ninfected_inperiod = sim_Binomial(S, p)
    End Function
    Public Function sim_Binomial(ByVal N As Double, ByVal p As Double)
        Dim temp As Object
        If N > 0 And p > 0 Then
            temp = myconnector.Evaluate("rbinom(1," & N & "," & p & ")")
            sim_Binomial = temp
        Else
            sim_Binomial = 0
        End If

    End Function
    Public Sub calc_nperiod()
        nperiods = Math.Round((ndaystorun * 24) / tau)
    End Sub

    Public Sub initiate_data_arrays()
        ReDim output_nsusceptible(nperiods)
        ReDim output_nlatent(nperiods)
        ReDim output_ninfectious(nperiods)
        ReDim output_nremoved(nperiods)
        ReDim output_nsubclincal(nperiods)
        ReDim output_nclincalinfectious(nperiods)
        ReDim Output_nclinical_removed(nperiods)
        ReDim Output_egg_rate_inperiod(nperiods)

    End Sub
    Public Sub writetoexcel(ByVal iter)
        Dim tempstartdate, tempenddate As Date
        tempstartdate = DateAndTime.Now
        Dim mytimearray(nperiods) As Double
        For i = 0 To nperiods
            mytimearray(i) = (i * tau) / 24
        Next
        myworksheet = myworkbook.Worksheets("Sus")
        myworksheet.Range(myworksheet.Cells(2 + iter, 2), myworksheet.Cells(2 + iter, 2 + nperiods + 1)).Value = (output_nsusceptible)
        myworksheet.Range(myworksheet.Cells(1, 2), myworksheet.Cells(1, 2 + nperiods + 1)).Value = (mytimearray)
        myworksheet.Cells(2 + iter, 1) = iter

        myworksheet = myworkbook.Worksheets("Latent")
        myworksheet.Range(myworksheet.Cells(2 + iter, 2), myworksheet.Cells(2 + iter, 2 + nperiods + 1)).Value = (output_nlatent)
        myworksheet.Range(myworksheet.Cells(1, 2), myworksheet.Cells(1, 2 + nperiods + 1)).Value = (mytimearray)
        myworksheet.Cells(2 + iter, 1) = iter

        myworksheet = myworkbook.Worksheets("Infectious")
        myworksheet.Range(myworksheet.Cells(2 + iter, 2), myworksheet.Cells(2 + iter, 2 + nperiods + 1)).Value = (output_ninfectious)
        myworksheet.Range(myworksheet.Cells(1, 2), myworksheet.Cells(1, 2 + nperiods + 1)).Value = (mytimearray)
        myworksheet.Cells(2 + iter, 1) = iter

        If True Then
            myworksheet = myworkbook.Worksheets("SubClinical")
            myworksheet.Range(myworksheet.Cells(2 + iter, 2), myworksheet.Cells(2 + iter, 2 + nperiods + 1)).Value = (output_nsubclincal)
            myworksheet.Range(myworksheet.Cells(1, 2), myworksheet.Cells(1, 2 + nperiods + 1)).Value = (mytimearray)
            myworksheet.Cells(2 + iter, 1) = iter


            myworksheet = myworkbook.Worksheets("Clinical infec")
            myworksheet.Range(myworksheet.Cells(2 + iter, 2), myworksheet.Cells(2 + iter, 2 + nperiods + 1)).Value = (output_nclincalinfectious)
            myworksheet.Range(myworksheet.Cells(1, 2), myworksheet.Cells(1, 2 + nperiods + 1)).Value = (mytimearray)
            myworksheet.Cells(2 + iter, 1) = iter

        End If

        myworksheet = myworkbook.Worksheets("Rem")
        myworksheet.Range(myworksheet.Cells(2 + iter, 2), myworksheet.Cells(2 + iter, 2 + nperiods + 1)).Value = (output_nremoved)
        myworksheet.Range(myworksheet.Cells(1, 2), myworksheet.Cells(1, 2 + nperiods + 1)).Value = (mytimearray)
        ' myworksheet.Range(myworksheet.Cells(3, 88), myworksheet.Cells(3 + niterations - 1, 88)).Value = (flocksizearray)
        myworksheet.Cells(2 + iter, 1) = iter

        ' myworksheet = myworkbook.Worksheets("Milk")
        ' myworksheet.Range(myworksheet.Cells(2 + iter, 2), myworksheet.Cells(2 + iter, 2 + nperiods + 1)).Value = ""
        ' myworksheet.Range(myworksheet.Cells(1, 2), myworksheet.Cells(1, 2 + nperiods + 1)).Value = (mytimearray)
        ' myworksheet.Cells(2 + iter, 2) = iter


        'myworksheet = myworkbook.Worksheets("Egg")
        'myworksheet.Range(myworksheet.Cells(2 + iter, 2), myworksheet.Cells(2 + iter, 2 + nperiods + 1)).Value = Output_egg_rate_inperiod
        'myworksheet.Range(myworksheet.Cells(1, 2), myworksheet.Cells(1, 2 + nperiods + 1)).Value = (mytimearray)
        'myworksheet.Cells(2 + iter, 1) = iter

        tempenddate = DateAndTime.Now
        excelwritingtime = excelwritingtime.Add(tempenddate.Subtract(tempstartdate))






    End Sub

    Public Sub WRITEflocksize()
        Dim r As Integer
        myworksheet = myworkbook.Worksheets("Flocksize")
        For r = 1 To niterations
            myworksheet.Cells(2 + r, 2).Value = (flocksizearray(r - 1))
        Next
    End Sub
    Public Sub WRITEstarteggrate()
        myworksheet = myworkbook.Worksheets("Flocksize")
        Dim r As Integer
        For r = 1 To niterations
            myworksheet.Cells(2 + r, 5).Value = (starting_egg_rate_array(r - 1))
        Next
    End Sub

    Public Sub record_data_one_iteration(ByVal periodindex As Integer, ByVal current_iter As Integer)
        output_nsusceptible(periodindex - 1) = pens(1).nsusceptible
        output_nlatent(periodindex - 1) = pens(1).nlatent
        output_ninfectious(periodindex - 1) = pens(1).ninfectious
        output_nremoved(periodindex - 1) = pens(1).nremoved
        output_nsubclincal(periodindex - 1) = pens(1).ninfectious - pens(1).nclincalinfectious
        output_nclincalinfectious(periodindex - 1) = pens(1).nclincalinfectious
        Output_nclinical_removed(periodindex - 1) = pens(1).nclinical_removed
        Output_egg_rate_inperiod(periodindex - 1) = calculate_egg_production_rate(pens(1).ninfectious, pens(1).nremoved)
    End Sub
    Public Function getaverage(ByVal mydblarray() As Double) As Double
        Dim myreturn As Object
        Try
            myconnector.SetSymbol("zmeanz", mydblarray)
            myreturn = myconnector.Evaluate("mean(zmeanz)")
            getaverage = CDbl(myreturn)
        Catch ex As Exception
            getaverage = 0
        End Try

    End Function

    Public Sub iterate(ByVal curent_iter)
        current_time = 0
        Dim tempstartdate, tempenddate, templstartdate, templendate As Date
        tempcontactrateiter = contactratearray(curent_iter - 1)
        templstartdate = DateAndTime.Now
        Call loadtimedistributionsr()
        Call createpensandcows(curent_iter)
        Call infectacow(curent_iter)
        Call set_temp_egg_rate(curent_iter)
        Call initiate_data_arrays()
        templendate = DateAndTime.Now
        distributionloadingtime = distributionloadingtime.Add(templendate.Subtract(templstartdate))
        For Period_index = 1 To nperiods
            Call record_data_one_iteration(Period_index, curent_iter)
            tempstartdate = DateAndTime.Now
            current_time = current_time + tau
            Call pens(1).iterate(current_time)
            tempenddate = DateAndTime.Now
            barniteratingtime = barniteratingtime.Add(tempenddate.Subtract(tempstartdate))
        Next Period_index
    End Sub
    Public Sub transrunparamterrecorder()
        myworksheet = myworkbook.Worksheets("Runp")
        If is_turkey_hen Then
            myworksheet.Cells(3, 2) = "Live Turkeys hens"
        Else
            myworksheet.Cells(3, 2) = "Live Turkeys toms"
        End If
        myworksheet.Cells(2, 2) = niterations
        myworksheet.Cells(5, 2) = getaverage(contactratearray)
        myworksheet.Cells(8, 2) = getaverage(flocksizearray)
        myworksheet.Cells(6, 2) = getaverage(cowinftimearray)
        myworksheet.Cells(7, 2) = getaverage(cowlatenttimearray)
        myworksheet.Cells(9, 2) = starttime
        myworksheet.Cells(10, 2) = endtime
        myworksheet.Cells(11, 2) = DateAndTime.DateDiff(DateInterval.Second, starttime, endtime)
        If iscontactratestochastic Then
            myworksheet.Cells(12, 2) = "stoc contact rate"
        Else
            myworksheet.Cells(12, 2) = "det contact rate"
        End If
    End Sub



    Public Sub dummybutton2_click()
        'This is used when multirunning
        initiate_statconnector()
        Dim mysurvsim As survsim
        mysurvsim = New survsim(myconnector)
        'Code to change any survesim paramter
        mysurvsim.movement_day_lower = Me.survsimparametertochange
        Me.Givendiseasedayonmovment = Me.survsimparametertochange
        mysurvsim.main()
    End Sub






    Public Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        initiate_statconnector()
        Dim mysurvsim As survsim
        mysurvsim = New survsim(myconnector)
        mysurvsim.main()

    End Sub



    '' end copy
    ''' <summary>

    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        myspeech = New Speech.Synthesis.SpeechSynthesizer
        myspeech.SelectVoiceByHints(Speech.Synthesis.VoiceGender.Female, Speech.Synthesis.VoiceAge.Adult)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        starttime = DateAndTime.Now
        Call calc_nperiod()
        Call initate_excel()
        Call initiate_statconnector()
        Call loadstochasticarrays()
        loadflocksize()
        myspeech.SpeakAsync("Hi running transmission")
        For k = 1 To niterations
            nbirds = flocksizearray(k - 1)
            If Math.IEEERemainder(k, updateform) = 0 Then
                TextBox1.Text = k
                Me.Update()
            End If


            Call calc_nperiod()
            Call iterate(k)
            Call writetoexcel(k - 1)
        Next k
        WRITEflocksize()
        WRITEstarteggrate()
        endtime = DateAndTime.Now
        Call transrunparamterrecorder()
        myconnector.Close()
        Call close_excel()
        myspeech.SpeakAsync("Completed transmission")
    End Sub
End Class
