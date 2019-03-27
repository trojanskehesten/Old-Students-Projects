Attribute VB_Name = "Module1"
Sub Center_Finding()

' Center_Finding searches the centre of input cloud.
' Warning! Some calculations are done on the list of Excel.
' D. Beregovoi, 26-05-2016

col_MSE_cur = 9 'column of current Mean Square Error (MSE)
row_MSE_cur = 9 'row of current MSE
MSE = Cells(row_MSE_cur, col_MSE_cur) 'initial MSE
step = Cells(5, 9) 'step of 1 iteration
t = Cells(8, 9) ' number of iterations
row_Xcoord = 6 'row of initial (current) X0
row_Ycoord = 7 'row of initial (current) Y0
col_Result = 9 'column of output data

For i = 1 To t
' --------------------------------------------------------------
' Moving by X coordinate
MSE2 = Cells(row_MSE_cur, col_MSE_cur) 'remember the current MSE
Cells(row_Xcoord, col_Result) = Cells(row_Xcoord, col_Result) + step 'cells(row_Xcoord, col_Result) is current X coordinate
MSE = Cells(row_MSE_cur, col_MSE_cur)
'Cells(37, 16) is calculated on the list of Excel.
' It is a sum of square differences between current centre and position of all markers.
If MSE > MSE2 Then Cells(row_Xcoord, col_Result) = Cells(row_Xcoord, col_Result) - 2 * step
MSE = Cells(row_MSE_cur, col_MSE_cur)
If MSE > MSE2 Then Cells(row_Xcoord, col_Result) = Cells(row_Xcoord, col_Result) + step

' --------------------------------------------------------------
' Moving by Y coordinate
MSE2 = Cells(row_MSE_cur, col_MSE_cur)
Cells(row_Ycoord, col_Result) = Cells(row_Ycoord, col_Result) + step 'cells(row_Ycoord, col_Result) is current Y coordinate
MSE = Cells(row_MSE_cur, col_MSE_cur)
If MSE > MSE2 Then Cells(row_Ycoord, col_Result) = Cells(row_Ycoord, col_Result) - 2 * step
MSE = Cells(row_MSE_cur, col_MSE_cur)
If MSE > MSE2 Then Cells(row_Ycoord, col_Result) = Cells(row_Ycoord, col_Result) + step
Next i
End Sub
