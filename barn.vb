Public Class barn
    Public susceptible As ArrayList
    Public latent As ArrayList
    Public infectious As ArrayList
    Public removed As ArrayList
    Public nsusceptible As Integer
    Public nlatent As Integer
    Public ninfectious As Integer
    Public nremoved As Integer
    Public nsubclincal As Integer
    Public nclincalinfectious As Integer
    Public nclinical_removed As Integer

    Public ncow As Integer
    Public capacity As Integer
    Public penid As Integer

    ' New variables
    Public new_susceptible As ArrayList
    Public new_latent As ArrayList
    Public new_infectious As ArrayList
    Public new_removed As ArrayList
    Public new_nsusceptible As Integer
    Public new_nlatent As Integer
    Public new_ninfectious As Integer
    Public new_nremoved As Integer
    Public new_nsubclincal As Integer
    Public new_nclincalinfectious As Integer
    Public new_nclinical_removed As Integer


    Dim ntest() As bird

    Public Sub pen_iterate()

    End Sub
    Public Sub pop_new_with_old()
        new_nsusceptible = nsusceptible
        new_nlatent = nlatent
        new_ninfectious = ninfectious
        new_nremoved = nremoved
        new_nsubclincal = nsubclincal
        new_nclincalinfectious = 0
        new_nclinical_removed = 0
    End Sub
    Public Sub iterate(ByVal curtime As Single)
        pop_new_with_old()
        infect_to_remove(curtime)
        latent_to_infectious(curtime)
        Sus_to_latent(curtime)
        check_infectious_for_clinical(curtime)
        check_removed_for_clinical(curtime)
        pop_old_with_new()
    End Sub
    Public Sub pop_old_with_new()
        nsusceptible = new_nsusceptible
        nlatent = new_nlatent
        ninfectious = new_ninfectious
        nremoved = new_nremoved
        nsubclincal = new_nsubclincal
        nclincalinfectious = new_nclincalinfectious
        nclinical_removed = new_nclinical_removed
    End Sub
    Public Sub infect_to_remove(ByVal curtime As Single)
        Dim i As Integer
        Dim k As Integer
        k = 0
        For i = 0 To infectious.Count - 1
            If infectious(i - k).lengthinfectious + infectious(i - k).timeofinfectious <= curtime Then
                infectious(i - k).timeofremoved = curtime
                removed.Add(infectious(i - k))
                new_nremoved += 1
                infectious.RemoveAt(i - k)
                k += 1
                new_ninfectious -= 1
            End If
        Next
    End Sub
    Public Sub latent_to_infectious(ByVal curtime As Single)
        Dim i As Integer
        Dim k As Integer
        k = 0
        For i = 0 To latent.Count - 1
            If latent(i - k).Lengthlat + latent(i - k).timeofinfection <= curtime Then
                latent(i - k).timeofinfectious = curtime
                infectious.Add(latent(i - k))
                new_ninfectious += 1
                new_nlatent -= 1
                latent.RemoveAt(i - k)
                k += 1
            End If
        Next
    End Sub
    Public Sub Sus_to_latent(ByVal curtime As Single)
        Dim i As Integer

        Dim ncowsinfected As Long
        ncowsinfected = Form1.calc_ninfected_inperiod(nsusceptible, ninfectious, Form1.nbirds, nremoved)
        For i = 0 To ncowsinfected - 1
            Dim tempcow As bird
            tempcow = susceptible(i)
            tempcow.timeofinfection = curtime
            latent.Add(tempcow)
        Next
        susceptible.RemoveRange(0, ncowsinfected)
        new_nsusceptible = new_nsusceptible - ncowsinfected
        new_nlatent = new_nlatent + ncowsinfected
    End Sub
    Public Sub check_infectious_for_clinical(ByVal curtime As Single)
        If True Then
            Dim testcow As bird
            For Each testcow In infectious
                If testcow.Lengthlat - (testcow.timeofinfectious) + curtime >= Form1.eggstarttime Then
                    If testcow.isclinical = False Then
                        testcow.timeofclinical = curtime
                        testcow.isclinical = True
                    End If
                    new_nclincalinfectious += 1
                End If
            Next
        End If
    End Sub
    Public Sub check_removed_for_clinical(ByVal curtime As Single)

    End Sub

    Public Sub addcow(ByVal mycow As bird)
        If capacity > ncow Then
            susceptible.Add(mycow)
        End If
    End Sub

End Class
