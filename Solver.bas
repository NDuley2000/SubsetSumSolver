Sub FindCombinationWithMinNumbersAndHighlightoptimised()
    Dim rng As Range, targetValue As Double
    Dim numArray() As Double, bestResult() As Boolean
    Dim i As Long, count As Long
    Dim cell As Range
    Dim minCount As Long

    ' Prompt user to select a range
    On Error Resume Next
    Set rng = Application.InputBox("Select the range of numbers:", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    ' Clear previous formatting
    rng.Interior.ColorIndex = xlNone

    ' Input target value
    targetValue = Application.InputBox("Enter the target value:", Type:=1)

    ' Store values in an array
    count = rng.Cells.count
    ReDim numArray(1 To count)
    ReDim bestResult(1 To count)

    i = 1
    For Each cell In rng
        numArray(i) = cell.Value
        i = i + 1
    Next cell

    ' Initialize minCount to a large value
    minCount = count + 1

    ' Call optimised function
    If FindCombinationOptimised(numArray, targetValue, count, bestResult, minCount) Then
        ' Highlight cells for the found combination
        i = 1
        For Each cell In rng
            If bestResult(i) Then cell.Interior.Color = RGB(144, 238, 144) ' Light green
            i = i + 1
        Next cell
        MsgBox "Combination found and highlighted!"
    Else
        MsgBox "No combination adds up to the target value."
    End If
End Sub

Function FindCombinationOptimised(numArray() As Double, targetValue As Double, count As Long, bestResult() As Boolean, ByRef minCount As Long) As Boolean
    Dim stack() As Variant
    Dim used() As Boolean
    Dim i As Long, currentSum As Double, currentCount As Long
    Dim stackPointer As Long
    Dim combination As Variant

    ' Initialise stack and used array
    ReDim stack(1 To 1)
    ReDim used(1 To count)
    stack(1) = Array(1, 0, 0, used) ' (Index, Current Sum, Current Count, Used Array)
    stackPointer = 1 ' Initialise stack pointer
    FindCombinationOptimised = False

    ' Process the stack iteratively
    Do While stackPointer > 0
        ' Retrieve the current state from the top of the stack
        combination = stack(stackPointer)
        stackPointer = stackPointer - 1 ' Pop the stack

        ' Retrieve current state
        i = combination(0)
        currentSum = combination(1)
        currentCount = combination(2)
        used = combination(3)

        ' Check if current sum matches the target
        If Abs(currentSum - targetValue) < 0.0000001 Then
            If currentCount < minCount Then
                minCount = currentCount
                bestResult = used
                FindCombinationOptimised = True
            End If
            GoTo NextIteration
        End If

        ' Skip invalid paths
        If i > count Then GoTo NextIteration

        ' Include the current number (direct modification of the "used" array)
        used(i) = True
        stackPointer = stackPointer + 1
        ReDim Preserve stack(1 To stackPointer)
        stack(stackPointer) = Array(i + 1, currentSum + numArray(i), currentCount + 1, used)

        ' Exclude the current number (backtrack by resetting the "used" array at this index)
        used(i) = False
        stackPointer = stackPointer + 1
        ReDim Preserve stack(1 To stackPointer)
        stack(stackPointer) = Array(i + 1, currentSum, currentCount, used)

NextIteration:
    Loop
End Function


