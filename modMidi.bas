Attribute VB_Name = "modMidi"
Option Explicit
Dim hMidiOutCopy As Long
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutReset Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Sub InitMidi()
Dim X As Long
Dim hMidiOut As Long
Dim MidiOk As Boolean
    X = midiOutOpen(hMidiOut, -1, 0&, 0&, 0&)
    If X = 0 Then
        hMidiOutCopy = hMidiOut
        MidiOk = True
    Else
        End
    End If
End Sub
Public Sub EndMidi()
Dim X As Integer
    X = midiOutClose(hMidiOutCopy)
End Sub
Private Function StopNote(Note, Channel)
Dim midiMsg As Long
    midiMsg = &H80 + (Note * &H100) + (Channel - 1)
    midiOutShortMsg hMidiOutCopy, midiMsg
End Function
Public Function PlayNote(Note As Long, Channel As Long, Velocity As Long, Durration As Long)
    Dim T As Long
    Dim CompleteMessage As Long
    Dim PartA As Long
    Dim PartB As Long
    T = GetTickCount
    'Work out what we are going to send.
    PartA = (Note * 256) + 143 + Channel
    PartB = (Velocity * 256) * 256
    CompleteMessage = PartA + PartB
    'Send the message.
    midiOutShortMsg hMidiOutCopy, CompleteMessage
    Do While GetTickCount <= T + Durration
    DoEvents
    Loop
    StopNote Note, Channel
End Function
Public Sub PauseNote(Durration As Long)
Dim T As Long
T = GetTickCount
    Do While GetTickCount <= T + Durration
    Loop
End Sub
