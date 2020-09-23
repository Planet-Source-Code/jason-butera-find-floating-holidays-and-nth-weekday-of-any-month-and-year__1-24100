Attribute VB_Name = "DateCalcs"
Option Explicit

Public Enum vbMonth
   vbJanuary = 1
   vbFebruary = 2
   vbMarch = 3
   vbApril = 4
   vbMay = 5
   vbJune = 6
   vbJuly = 7
   vbAugust = 8
   vbSeptember = 9
   vbOctober = 10
   vbNovember = 11
   vbDecember = 12
End Enum

Public Enum vbDayOccurrence
   vbFirst = 1
   vbSecond = 2
   vbThird = 3
   vbFourth = 4
   vbFifth = 5
   vbLast = 6
End Enum


Public Function GetDateByOccurrence( _
   iYear As Integer, _
   iMonth As vbMonth, _
   iWeekday As VbDayOfWeek, _
   iOccurrence As vbDayOccurrence) As Variant
   
   Dim intWeek As Integer
   Dim intWeekday As Integer
   Dim intDay As Integer
   Dim dtLastDay As Date
   Dim intLastDay As Integer
   Dim intDayTemp As Integer
      
   If iOccurrence = vbLast Then
      'GET LAST OCCURRENCE IN MONTH
      dtLastDay = DateSerial(iYear, iMonth + 1, 1 - 1)
      intLastDay = Day(dtLastDay)
      intWeekday = Weekday(dtLastDay) - 1
      intDay = intLastDay - (intWeekday - (iWeekday - 1))
   Else
      'GET SPECIFIED OCCURRENCE IN MONTH
      intWeek = 1 + ((iOccurrence - 1) * 7)
      intWeekday = Weekday(DateSerial(iYear, iMonth, intWeek))
      intDayTemp = iWeekday - intWeekday
      If intDayTemp < 0 Then
         intDayTemp = intDayTemp + 7
      End If
      intDay = (intWeek + intDayTemp)
   End If
   
   'CHECK TO SEE IF THERE IS NO Nth DAY OF THE MONTH
   If Not IsDate(iMonth & "/" & intDay & "/" & iYear) Then
      GetDateByOccurrence = False
   Else
      GetDateByOccurrence = DateSerial(iYear, iMonth, intDay)
   End If
   
End Function


Public Function IsHoliday(ByVal dtDate As Date) As Boolean

   Dim iYear As Integer
   Dim iMonth As vbMonth
   Dim iDay As Integer
   
   iYear = Year(dtDate)
   iMonth = Month(dtDate)
   iDay = Day(dtDate)

   'NEW YEARS DAY (JANUARY 1ST)
   If (iMonth = vbJanuary) And (iDay = 1) Then
      IsHoliday = True
      Exit Function
   End If
   
   'MLK B-DAY (3RD MONDAY IN JANUARY)
   If dtDate = GetDateByOccurrence(iYear, vbJanuary, vbMonday, 3) Then
      IsHoliday = True
      Exit Function
   End If
   
   'WASHINGTON B-DAY (3RD MONDAY IN FEBRUARY)
   If dtDate = GetDateByOccurrence(iYear, vbFebruary, vbMonday, 3) Then
      IsHoliday = True
      Exit Function
   End If
   
   'MEMORIAL DAY (LAST MONDAY IN MAY)
   If dtDate = GetDateByOccurrence(iYear, vbMay, vbMonday, vbLast) Then
      IsHoliday = True
      Exit Function
   End If
   
   '4th OF JULY (JULY 4TH)
   If (iMonth = vbJuly) And (iDay = 4) Then
      IsHoliday = True
      Exit Function
   End If
   
   'LABOR DAY (1ST MONDAY IN SEPTEMBER)
   If dtDate = GetDateByOccurrence(iYear, vbSeptember, vbMonday, 1) Then
      IsHoliday = True
      Exit Function
   End If
   
   'COLOMBUS DAY (2ND MONDAY IN OCTOBER)
   If dtDate = GetDateByOccurrence(iYear, vbOctober, vbMonday, 2) Then
      IsHoliday = True
      Exit Function
   End If

   'VETERANS DAY (NOVEMBER 11TH)
   If (iMonth = vbNovember) And (iDay = 11) Then
      IsHoliday = True
      Exit Function
   End If
   
   'THANKSGIVING (4TH THURSDAY IN NOVEMBER)
   If dtDate = GetDateByOccurrence(iYear, vbNovember, vbThursday, 4) Then
      IsHoliday = True
      Exit Function
   End If
   
   'CHRISTMAS (DECEMBER 25TH)
   If (iMonth = vbDecember) And (iDay = 25) Then
      IsHoliday = True
      Exit Function
   End If
End Function

