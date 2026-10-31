Attribute VB_Name = "AddNarration"
' ------------------------------------------------------------------
' Inserts audio\slide01.mp3, slide02.mp3, ... into the matching slides,
' makes each one play automatically (icon hidden), sets
' "Advance after" = audio length + PAD_SECONDS, and turns any
' on-click animations into automatic ones so the video never stalls.
'
' FALLBACK only, if PowerPoint rejects export/*_mit_audio.pptx: open the ORIGINAL
' pptx in videos/<video>/ (next to the audio\ folder) and run this macro.
' Run it again at any time: old narration shapes are replaced.
' Then: File > Export > Create a Video > "Use Recorded Timings and Narrations"
' ------------------------------------------------------------------
Const PAD_SECONDS As Single = 1.5         ' silence after the voice before the next slide
Const ANIM_DELAY_SECONDS As Single = 1   ' when former on-click animations appear

Sub AddNarrationToAllSlides()
    Dim folder As String, f As String, sld As Slide, shp As Shape
    Dim i As Long, j As Long, secs As Single, eff As Effect, report As String

    folder = ActivePresentation.Path & Application.PathSeparator & "audio" & Application.PathSeparator

    For Each sld In ActivePresentation.Slides
        ' remove narration from an earlier run
        For j = sld.Shapes.Count To 1 Step -1
            If sld.Shapes(j).Name = "Narration" Then sld.Shapes(j).Delete
        Next j

        ' existing on-click animations -> start automatically ANIM_DELAY_SECONDS into the slide
        For j = 1 To sld.TimeLine.MainSequence.Count
            Set eff = sld.TimeLine.MainSequence(j)
            If eff.Timing.TriggerType = msoAnimTriggerOnPageClick Then
                eff.Timing.TriggerType = msoAnimTriggerWithPrevious
                eff.Timing.TriggerDelayTime = eff.Timing.TriggerDelayTime + ANIM_DELAY_SECONDS
            End If
        Next j

        f = folder & "slide" & Format(sld.SlideIndex, "00") & ".mp3"
        If Dir(f) = "" Then
            report = report & "Slide " & sld.SlideIndex & ": NO AUDIO FILE" & vbCrLf
        Else
            Set shp = sld.Shapes.AddMediaObject2(f, msoFalse, msoTrue, _
                        ActivePresentation.PageSetup.SlideWidth - 40, 10, 30, 30)
            shp.Name = "Narration"

            ' play with the slide, before any other animation, icon hidden
            Set eff = sld.TimeLine.MainSequence.AddEffect(shp, msoAnimEffectMediaPlay, , msoAnimTriggerWithPrevious)
            eff.MoveTo 1
            shp.AnimationSettings.PlaySettings.HideWhileNotPlaying = msoTrue

            secs = shp.MediaFormat.Length / 1000
            With sld.SlideShowTransition
                .AdvanceOnClick = msoFalse
                .AdvanceOnTime = msoTrue
                .AdvanceTime = secs + PAD_SECONDS
            End With
            report = report & "Slide " & sld.SlideIndex & ": " & Format(secs, "0.0") & " s audio -> advances after " & _
                     Format(secs + PAD_SECONDS, "0.0") & " s" & vbCrLf
        End If
    Next sld

    MsgBox report, vbInformation, "Narration added"
End Sub
