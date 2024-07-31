VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Ej samma
Public Function Laps(ByVal Ver As GP2Ver) As Variant
    If Ver = FR Then
        Laps = 1951897
    ElseIf Ver = IT Then
        Laps = 1951901
    ElseIf Ver = NL Then
        Laps = 1951897
    ElseIf Ver = SP Then
        Laps = 1951933
    ElseIf Ver = TY Then
        Laps = 1951961
    ElseIf Ver = UK Then
        Laps = 1951849
    ElseIf Ver = US Then
        Laps = 1951849
    End If
End Function

Public Function Text(ByVal Ver As GP2Ver) As Variant
    If Ver = FR Then
        Text = 5686683
    ElseIf Ver = IT Then
        Text = 5685991
    ElseIf Ver = NL Then
        Text = 5684786
    ElseIf Ver = SP Then
        Text = 5687337
    ElseIf Ver = TY Then
        Text = 5686490
    ElseIf Ver = UK Then
        Text = 5682445
    ElseIf Ver = US Then
        Text = 5682444
    End If
End Function

Public Function Point(ByVal Ver As GP2Ver) As Variant
    If Ver = FR Then
        Point = 1954110
    ElseIf Ver = IT Then
        Point = 1954114
    ElseIf Ver = NL Then
        Point = 1954110
    ElseIf Ver = SP Then
        Point = 1954146
    ElseIf Ver = TY Then
        Point = 1954174
    ElseIf Ver = UK Then
        Point = 1954062
    ElseIf Ver = US Then
        Point = 1954062
    End If
End Function
Public Function Length(ByVal Ver As GP2Ver) As Variant
    If Ver = FR Then
        Length = 1953886
    ElseIf Ver = IT Then
        Length = 1953890
    ElseIf Ver = NL Then
        Length = 1953886
    ElseIf Ver = SP Then
        Length = 1953922
    ElseIf Ver = TY Then
        Length = 1953950
    ElseIf Ver = UK Then
        Length = 1953838
    ElseIf Ver = US Then
        Length = 1953838
    End If
End Function
Public Function Level(ByVal Ver As GP2Ver) As Variant
    If Ver = FR Then
        Level = 5661835
    ElseIf Ver = IT Then
        Level = 5661632
    ElseIf Ver = NL Then
        Level = 5661382
    ElseIf Ver = SP Then
        Level = 5671247
    ElseIf Ver = TY Then
        Level = 5661503
    ElseIf Ver = UK Then
        Level = 5661257
    ElseIf Ver = US Then
        Level = 5661257
    End If
End Function
Public Function OneAsNull(ByVal Ver As GP2Ver) As Variant
    If Ver = FR Then
        OneAsNull = 1951896
    ElseIf Ver = IT Then
        OneAsNull = 1951900
    ElseIf Ver = NL Then
        OneAsNull = 1951896
    ElseIf Ver = SP Then
        OneAsNull = 1951932
    ElseIf Ver = TY Then
        OneAsNull = 1951960
    ElseIf Ver = UK Then
        OneAsNull = 1951848
    ElseIf Ver = US Then
        OneAsNull = 1951848
    End If
End Function
'Samma
Public Function Help() As Variant
    Help = 1224469
End Function
Public Function Ware() As Variant
    Ware = 1280585
End Function
Public Function SaveLapTime() As Variant
    SaveLapTime = 1029530
End Function
Public Function SaveLapTime2() As Variant
    SaveLapTime2 = 1078955
End Function
Public Function NoPitSpeed() As Variant
    NoPitSpeed = 596852
End Function
Public Function UseTeam() As Variant
    UseTeam = 678540
End Function
Public Function PQPower() As Variant
    PQPower = 678560
End Function
Public Function PRPower() As Variant
    PRPower = 678553
End Function
Public Function CWeight() As Variant
    CWeight = 1282253
End Function
Public Function PGrip() As Variant
    PGrip = 1282273
End Function
Public Function PWeight() As Variant
    PWeight = 1282249
End Function
Public Function PitSpeed() As Variant
    PitSpeed = 1280069
End Function
