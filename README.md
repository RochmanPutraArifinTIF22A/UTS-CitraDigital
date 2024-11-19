Dim h(255) As Integer ' Histogram for original image
Dim h2(255) As Integer ' Histogram for equalized image
Dim cdf(255) As Integer ' Cumulative Distribution Function (CDF)
Dim total_pixels As Integer
Dim r, g, b, xg, yg, w As Integer
Dim min_cdf As Integer

' Initialize histograms
For i = 0 To 255
    h(i) = 0
    h2(i) = 0
Next i

' Step 1: Calculate the histogram of the original image
For i = 0 To Picture1.ScaleWidth - 1
    For j = 0 To Picture1.ScaleHeight - 1
        ' Get the color value of the pixel
        w = Picture1.Point(i, j)
        
        ' Extract the RGB components
        r = w And RGB(255, 0, 0)
        g = Int((w And RGB(0, 255, 0)) / 256)
        b = Int((w And RGB(0, 0, 255)) / 256)
        
        ' Calculate grayscale value (average of RGB)
        xg = Int((r + g + b) / 3)
        
        ' Update the histogram for original image
        h(xg) = h(xg) + 1
    Next j
Next i

' Step 2: Calculate the cumulative distribution function (CDF)
total_pixels = Picture1.ScaleWidth * Picture1.ScaleHeight
cdf(0) = h(0)

For i = 1 To 255
    cdf(i) = cdf(i - 1) + h(i)
Next i

' Step 3: Apply histogram equalization to the image
min_cdf = cdf(0)
If min_cdf = 0 Then min_cdf = 1 ' Avoid division by zero

For i = 0 To Picture1.ScaleWidth - 1
    For j = 0 To Picture1.ScaleHeight - 1
        ' Get the color value of the pixel
        w = Picture1.Point(i, j)
        
        ' Extract the RGB components
        r = w And RGB(255, 0, 0)
        g = Int((w And RGB(0, 255, 0)) / 256)
        b = Int((w And RGB(0, 0, 255)) / 256)
        
        ' Calculate grayscale value (average of RGB)
        xg = Int((r + g + b) / 3)
        
        ' Apply the CDF-based equalization
        yg = Int((cdf(xg) - cdf(0)) * 255 / (total_pixels - min_cdf))
        
        ' Update the histogram for the equalized image
        h2(yg) = h2(yg) + 1
        
        ' Set the new grayscale pixel in Picture2
        Picture2.PSet(i, j), RGB(yg, yg, yg)
    Next j
Next i

' Step 4: Plot the histograms using MSChart
For i = 0 To 255
    MSChart1.Row = i + 1
    MSChart1.Data = h(i) ' Original histogram
    MSChart1.RowLabel = Trim(Str(i))
    
    MSChart2.Row = i + 1
    MSChart2.Data = h2(i) ' Equalized histogram
    MSChart2.RowLabel = Trim(Str(i))
Next i
