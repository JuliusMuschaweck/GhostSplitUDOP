Imports LTUDSimulation
Imports System.Math
Public Class GhostSplitUDOP
    Implements LTUDSimulation.IOpticalProperty
    Implements LTUDSimulation.IPropertySupport
    Implements LTUDSimulation.ISimulationSupportGeneral

    Private refractedPower_ As Double = 1
    Private reflectedPower_ As Double = 0
    Private absorbedPower_ As Double = 0
    Private refractedProb_ As Double = 1
    Private reflectedProb_ As Double = 0
    Private absorbedProb_ As Double = 0
    Private ReadOnly nNum_ As Integer = 6
    Private ReadOnly ranSeed As Integer = 42
    Private rand As System.Random = New System.Random(ranSeed) ' using a fixed seed value
    Private ranLock As New Object()
    Public Enum ErrorCodes As Integer
        ErrOk = 0
        ErrIllegalValue = -1
        ErrIllegalIndex = -2
        ErrInternalError = -3
    End Enum

    Private Sub Adjust(v1 As Double, ByRef v2 As Double, ByRef v3 As Double)
        ' Given v1, adjusts v2 and v3 such that sum = 1 and all >= 0
        ' If possible (i.e. v1+v2 <= 1), modifies only v3.
        If v1 + v2 + v3 = 1.0 Then
            Return
        ElseIf v1 + v2 + v3 < 1.0 Then
            v3 = 1.0 - v1 - v2
        Else ' v1+v2+v3> 1
            If v1 + v2 < 1.0 Then
                v3 = 1.0 - v1 - v2
            ElseIf v1 + v2 > 1.0 Then
                v2 = 1.0 - v1
                v3 = 0
            Else ' v1+v2= 1
                v3 = 0
            End If
        End If
    End Sub

    Public Function bendRay(prevRefInd As Double, afterRefInd As Double, isProbSplittingOn As LTBoolean, isPolTraceOn As LTBoolean, ByRef rays As LTOpticalPropertyRayData) As Integer Implements IOpticalProperty.bendRay
        Dim L, M, N As Double
        Dim doTIR As Boolean
        Dim aux As Double
        Dim r As Double
        Dim p As Double
        Dim tiny As Double = 0.000000000001
        Dim ireflect As Integer = -1
        Dim irefract As Integer = -1
        Dim nr As Integer = 0
        If absorbedPower_ = 1 Then
            rays.setNumberOfOutRays(0)
            Return ErrorCodes.ErrOk
        End If
        rays.getEnteringRayDirectionCos(L, M, N)
        aux = L * L + M * M - afterRefInd * afterRefInd
        doTIR = (aux >= 0)
        If doTIR Then ' reflect with 100% - absorption
            rays.setNumberOfOutRays(1)
            rays.setOutRayDirectionCos(0, L, M, -N)
            rays.setOutRayTransmittance(0, 1 - absorbedPower_)
        Else ' no TIR n-> split
            If isProbSplittingOn Then 'random reflect/refract
                SyncLock ranLock
                    r = rand.NextDouble()
                End SyncLock
                If r < refractedProb_ Then ' refract
                    p = refractedPower_ / (refractedProb_ + tiny)
                    If p < 0.1 * refractedPower_ / tiny Then
                        rays.setNumberOfOutRays(1)
                        rays.setOutRayDirectionCos(0, L, M, Math.Sqrt(-aux) * Math.Sign(N))
                        rays.setOutRayTransmittance(0, p)
                    Else ' power too small, do not trace
                        rays.setNumberOfOutRays(0)
                    End If
                ElseIf r < refractedProb_ + reflectedProb_ Then ' reflect
                    p = reflectedPower_ / (reflectedProb_ + tiny)
                    If p < 0.1 * reflectedPower_ / tiny Then
                        rays.setNumberOfOutRays(1)
                        rays.setOutRayDirectionCos(0, L, M, -N)
                        rays.setOutRayTransmittance(0, p)
                    Else ' power too small, do not trace
                        rays.setNumberOfOutRays(0)
                    End If
                Else ' absorb
                    rays.setNumberOfOutRays(0)
                End If
            Else 'split rays into two
                ' assign ray numbers to reflect/refract path
                If reflectedPower_ > 0 Then
                    ireflect = 0
                    nr += 1
                End If
                If refractedPower_ > 0 Then
                    irefract = ireflect + 1 ' 0 uf ireflect = -1
                    nr += 1
                End If
                rays.setNumberOfOutRays(nr)
                If ireflect >= 0 Then
                    rays.setOutRayDirectionCos(ireflect, L, M, -N)
                    rays.setOutRayTransmittance(ireflect, reflectedPower_)
                End If
                If irefract >= 0 Then
                    rays.setOutRayDirectionCos(irefract, L, M, Math.Sqrt(-aux) * Math.Sign(N))
                    rays.setOutRayTransmittance(irefract, refractedPower_)
                End If
            End If
        End If
        bendRay = ErrorCodes.ErrOk
    End Function

    Public Function getDescription() As String Implements IPropertySupport.getDescription
        getDescription = "GhostSplitUDOP: Like probabilistic split, but reflect/transmit/absorb fractions and probabilities separately adjustable. Created by Julius Muschaweck (www.jmoptics.de)"
    End Function

    Public Function getNumberOfNumericalAttributes() As Integer Implements IPropertySupport.getNumberOfNumericalAttributes
        Return 6
    End Function

    Public Function getValueOfNumericalAttribute(index As Integer) As Double Implements IPropertySupport.getValueOfNumericalAttribute
        Select Case index
            Case 0
                Return refractedPower_ * 100
            Case 1
                Return reflectedPower_ * 100
            Case 2
                Return absorbedPower_ * 100
            Case 3
                Return refractedProb_ * 100
            Case 4
                Return reflectedProb_ * 100
            Case 5
                Return absorbedProb_ * 100
            Case Else
                Return -1
        End Select
    End Function

    Public Function setValueOfNumericalAttribute(index As Integer, value As Double) As Integer Implements IPropertySupport.setValueOfNumericalAttribute
        If (value < 0) Or (value > 100) Then
            Return ErrorCodes.ErrIllegalValue
        End If
        Select Case index
            Case 0
                refractedPower_ = value * 0.01
                Adjust(refractedPower_, reflectedPower_, absorbedPower_)
            Case 1
                reflectedPower_ = value * 0.01
                Adjust(reflectedPower_, refractedPower_, absorbedPower_)
            Case 2
                absorbedPower_ = value * 0.01
                Adjust(absorbedPower_, refractedPower_, reflectedPower_)
            Case 3
                refractedProb_ = value * 0.01
                Adjust(refractedProb_, reflectedProb_, absorbedProb_)
            Case 4
                reflectedProb_ = value * 0.01
                Adjust(reflectedProb_, refractedProb_, absorbedProb_)
            Case 5
                absorbedProb_ = value * 0.01
                Adjust(absorbedProb_, refractedProb_, reflectedProb_)
            Case Else
                Return ErrorCodes.ErrIllegalIndex
        End Select
        Return ErrorCodes.ErrOk
    End Function

    Public Function getNameOfNumericalAttribute(index As Integer) As String Implements IPropertySupport.getNameOfNumericalAttribute
        Select Case index
            Case 0
                Return "refracted power (%)"
            Case 1
                Return "reflected power (%)"
            Case 2
                Return "absorbed power (%)"
            Case 3
                Return "refraction probability (%)"
            Case 4
                Return "reflection probability (%)"
            Case 5
                Return "termination probability (%)"
            Case Else
                Return "bad index " + index.ToString()
        End Select
    End Function

    Public Function getNumberOfStringAttributes() As Integer Implements IPropertySupport.getNumberOfStringAttributes
        getNumberOfStringAttributes = 0
    End Function

    Public Function getNameOfStringAttribute(index As Integer) As String Implements IPropertySupport.getNameOfStringAttribute
        getNameOfStringAttribute = "getNameOfStringAttribute: cannot happen"
    End Function

    Public Function getValueOfStringAttribute(index As Integer) As String Implements IPropertySupport.getValueOfStringAttribute
        getValueOfStringAttribute = "getValueOfStringAttribute: cannot happen"
    End Function

    Public Function setValueOfStringAttribute(index As Integer, strValue As String) As Integer Implements IPropertySupport.setValueOfStringAttribute
        setValueOfStringAttribute = ErrorCodes.ErrIllegalIndex
    End Function

    Public Function updateDefinition(iLTAPI As Object, zoneName As String) As Integer Implements IPropertySupport.updateDefinition
        updateDefinition = ErrorCodes.ErrOk ' nothing to do
    End Function

    Public Function getDefaultValueOfNumericalAttribute(index As Integer) As Double Implements IPropertySupport.getDefaultValueOfNumericalAttribute
        Select Case index
            Case 0
                Return 100
            Case 1
                Return 0
            Case 2
                Return 0
            Case 3
                Return 100
            Case 4
                Return 0
            Case 5
                Return 0
            Case Else
                Return -1
        End Select
    End Function

    Public Function getDefaultValueOfStringAttribute(index As Integer) As String Implements IPropertySupport.getDefaultValueOfStringAttribute
        getDefaultValueOfStringAttribute = "getDefaultValueOfStringAttribute: cannot happen"
    End Function

    Public Function getErrorText(errorNum As Integer) As String Implements IPropertySupport.getErrorText
        Select Case errorNum
            Case ErrorCodes.ErrOk
                Return "OK"
            Case ErrorCodes.ErrIllegalValue
                Return "Illegal value"
            Case ErrorCodes.ErrIllegalIndex
                Return "Illegal index"
            Case ErrorCodes.ErrInternalError
                Return "Internal error"
            Case Else
                Return "Unknown error code"
        End Select
    End Function

    Public Function onEvent(iLTAPI As Object, theEvent As LTNotification) As Integer Implements IPropertySupport.onEvent
        onEvent = ErrorCodes.ErrOk
    End Function

    Public Function doUserButton(iLTAPI As Object, buttonNumber As Integer) As Integer Implements IPropertySupport.doUserButton
        Select Case buttonNumber
            Case 1
                rand = New System.Random(ranSeed)
                Dim lt As LightTools.ILTAPI4
                lt = iLTAPI
                lt.Message("GhostSplitUDOP: resetting internal random number generator")
            Case 2
            Case 3
            Case 4
            Case Else
        End Select
        doUserButton = ErrorCodes.ErrOk
    End Function

    Public Function getUserButtonLabel(buttonNumber As Integer) As String Implements IPropertySupport.getUserButtonLabel
        Select Case buttonNumber
            Case 1
                Return "Reset RanGen"
            Case 2
                Return "NoOp"
            Case 3
                Return "NoOp"
            Case 4
                Return "NoOp"
            Case Else
                Return "Button " + buttonNumber.ToString()
        End Select
    End Function

    Public Function beginSimulation(iLTAPI As Object, beginStatus As LTSimulationBeginType) As Integer Implements ISimulationSupportGeneral.beginSimulation
        beginSimulation = ErrorCodes.ErrOk
    End Function

    Public Sub finishSimulationInterval(iLTAPI As Object, finishStatus As LTSimulationIntervalType) Implements ISimulationSupportGeneral.finishSimulationInterval
    End Sub

    Public Sub beginNSRayTrace(iLTAPI As Object) Implements ISimulationSupportGeneral.beginNSRayTrace
    End Sub

    Public Sub finishNSRayTrace(iLTAPI As Object) Implements ISimulationSupportGeneral.finishNSRayTrace
    End Sub

    Public Sub beginRayTrace(rayOrdinalNumber As ULong, wavelength As Double, power As Double, powerUnit As LTPowerUnits) Implements ISimulationSupportGeneral.beginRayTrace
    End Sub

    Public Sub endRayTrace(rayOrdinalNumber As ULong, status As LTBoolean) Implements ISimulationSupportGeneral.endRayTrace
    End Sub

    Public Function resetRandomSeeds(iLTAPI As Object) As LTBoolean Implements ISimulationSupportGeneral.resetRandomSeeds
        resetRandomSeeds = LTBoolean.LT_False
    End Function

    Public Sub postSimulation(iLTAPI As Object, ByRef userData As Array) Implements ISimulationSupportGeneral.postSimulation
    End Sub

    Public Function isThreadSafe() As LTBoolean Implements ISimulationSupportGeneral.isThreadSafe
        isThreadSafe = LTBoolean.LT_True
    End Function

    Public Function getPolarizationCoordSys() As LTPolarizeCoordSysType Implements ISimulationSupportGeneral.getPolarizationCoordSys
        getPolarizationCoordSys = LTPolarizeCoordSysType.LT_POL_COORD_SP
    End Function

    Public Function getCoordType(iLTAPI As Object) As LTCoordinateType Implements ISimulationSupportGeneral.getCoordType
        getCoordType = LTCoordinateType.LT_COORDINATE_ENTITY
    End Function

    Public Sub enterEntity(entityName As String) Implements ISimulationSupportGeneral.enterEntity
    End Sub

    Public Sub exitEntity(entityName As String) Implements ISimulationSupportGeneral.exitEntity
    End Sub
End Class
