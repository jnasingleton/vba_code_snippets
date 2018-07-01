Public Function Haversine(ByVal Lat1 As Double, ByVal Lon1 As Double, ByVal Lat2 As Double, ByVal Lon2 As Double)
'Source: https://stackoverflow.com/questions/35175057/vba-haversine-formula
'Credit: https://stackoverflow.com/users/5876361/banshe, https://stackoverflow.com/users/3894917/tom-sharpe
' This function returns the Haversine distance between two points
Dim R As Integer, 
Dim dlon As Double, dlat As Double, Rad1 As Double, Rad2 As Double
Dim a As Double, c As Double, d As Double

' Returns the distance (in km) between (Lat1, Lon1) and (Lat2, Lon2) points

R = 6371

dlon = Excel.WorksheetFunction.Radians(Lon2 - Lon1)
dlat = Excel.WorksheetFunction.Radians(Lat2 - Lat1)
Rad1 = Excel.WorksheetFunction.Radians(Lat1)
Rad2 = Excel.WorksheetFunction.Radians(Lat2)

a = Sin(dlat / 2) * Sin(dlat / 2) + Cos(Rad1) * Cos(Rad2) * Sin(dlon / 2) * Sin(dlon / 2)
c = 2 * Excel.WorksheetFunction.Atan2(Sqr(1 - a), Sqr(a))
d = R * c

Haversine = d

End Function