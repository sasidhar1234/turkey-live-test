Public Class bird

    Public isinfected As Boolean
    Public Penid As Integer
    Public cowstate As String
    Public Lengthlat As Single
    Public lengthsubclin As Single
    Public lengthinfectious As Single
    Public lengthclinical As Single
    Public cowid As Integer
    Public timeofinfection As Single = 10000
    Public timeofinfectious As Single = 10000
    Public timeofclinical As Single = 10000
    Public timeofremoved As Single = 10000
    Public issubclinical As Boolean
    Public isclinical As Boolean







    Public Sub New(ByVal pen As Integer, ByVal cowiden As Integer)
        isinfected = False
        cowstate = "S"
        Penid = pen
        cowid = cowiden
        Lengthlat = -1
        lengthinfectious = -1
        lengthsubclin = -1
        issubclinical = False
        isclinical = False
    End Sub
    Public Function getfmdstatus() As Boolean
        getfmdstatus = isinfected
    End Function




End Class
