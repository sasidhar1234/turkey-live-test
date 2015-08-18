Public Class Mortalitydatasimulator
    Dim myconnector As STATCONNECTORSRVLib.StatConnector
    Dim counter As Long
    Public ndays1 As Integer
    Public Sub New(ByVal Ndays As Integer, ByVal con As STATCONNECTORSRVLib.StatConnector)
        myconnector = con
        ndays1 = Ndays
    End Sub
    Public Sub getcontinousmortdata(ByVal newflocksize As Long, ByRef countarray() As Double)
        Dim temp_flocksize As Double

        temp_flocksize = newflocksize
        myconnector.EvaluateNoReturn("flockno<-sample.int(32,size = 1)")
        myconnector.EvaluateNoReturn("endpoint<-sample.int((3),size =1)")
        myconnector.EvaluateNoReturn("newflocksize<-" & temp_flocksize)
        ''myconnector.EvaluateNoReturn("estflocksize<-weeklyavcount[flockno]/sample(weakflockarr[,1],size=1)")
        myconnector.EvaluateNoReturn("selectedcounts<-allflocks[(maxn-endpoint-13):(maxn-endpoint),flockno]")
        myconnector.EvaluateNoReturn("weakflockrate<-rlnorm(1, -5.0369,.3330)")
        '0.0056679*(rbeta(10000, 1.9073,2.0001))
        myconnector.EvaluateNoReturn("curweekmor<-sum(selectedcounts[8:14])")
        counter += 1
        countarray = myconnector.Evaluate("round(newflocksize*selectedcounts*weakflockrate/curweekmor)")


    End Sub
    Public Sub setuprarrays()
        myconnector.EvaluateNoReturn("allflocks <- read.csv(""C:\\Codes\\Live turkey Movement\\mortality\\Dailymort1.csv"",header=TRUE)")
        ' myconnector.EvaluateNoReturn("maxn2 <- read.csv(""C:\\Codes\\Live turkey Movement\\mortality\\maxn.csv"",header=TRUE)")
        Dim temp_weeklymortality As Double
        myconnector.EvaluateNoReturn("maxn<-21")
        myconnector.EvaluateNoReturn("selectedcounts<-1")
        ' myconnector.EvaluateNoReturn("weeklyavcount<-1:12")
        'myconnector.EvaluateNoReturn("for (i in 1:12) weeklyavcount[i]<-sum(allflocks[i],na.rm=TRUE)/(maxn[i]/7)")
        myconnector.EvaluateNoReturn("flockno<-1")
        myconnector.EvaluateNoReturn("endpoint<-1")
        myconnector.EvaluateNoReturn("length<-" & ndays1)
    End Sub

End Class
