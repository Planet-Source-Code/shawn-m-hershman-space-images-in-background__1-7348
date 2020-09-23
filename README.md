<div align="center">

## Space Images in Background


</div>

### Description

Evenly spaces an image as many times as will fit on a form
 
### More Info
 
1) Create a picture in the form and set it's INDEX property to 0 and it's VISIBLE property to false

2) SHOW the form before calling this subroutine

None I'm aware of


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Shawn M\. Hershman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shawn-m-hershman.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shawn-m-hershman-space-images-in-background__1-7348/archive/master.zip)





### Source Code

```
Private Sub Space_Images()
Dim PicCols As Integer
Dim PicRows As Integer
Dim HExtraSpace As Integer
Dim VExtraSpace As Integer
Dim HSpacing As Integer
Dim VSpacing As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
On Error Resume Next
 'Calculate the appropriate spacings
 PicCols = CInt((Me.Width / picPicture(0).Width) - 0.5)
 PicRows = CInt((Me.Height / picPicture(0).Height) - 0.5)
 HExtraSpace = Me.Width - (picPicture(0).Width * PicCols)
 VExtraSpace = Me.Height - (picPicture(0).Height * PicRows)
 HSpacing = CInt((HExtraSpace / (PicCols + 1)) - 0.5)
 VSpacing = CInt((VExtraSpace / (PicRows + 1)) - 0.5)
 'Display the background images
 For i = 0 To PicRows - 1
 For j = 0 To PicCols - 1
  k = (PicCols * i) + j
  Load picPicture(k)
  picPicture(k).Left = (HSpacing * (j + 1)) + (picPicture(0).Width * j)
  picPicture(k).Top = (VSpacing * (i + 1)) + (picPicture(0).Height * i)
  picPicture(k).Visible = True
 Next j
 Next i
End Sub
```

