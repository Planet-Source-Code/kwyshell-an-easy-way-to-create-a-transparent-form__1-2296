<div align="center">

## An Easy Way to Create a Transparent Form \!


</div>

### Description

I provide an easy to create the non-rectangle form with usercontrol. We may use the control to design a non-rectange form or a desktop animation easily.
 
### More Info
 
Sometime we will want to create a non-rectangle window to make our UI so Cool! But the only way to create a non-rectangle windows is just using the windows regions API.We can use the SetWindowRgn to change the window region. So to create a non-rectangle region is only thing we have to do and it's also the only problem we need to resolve.

How to create an non-rectangle region ? The only way I know to do this is using the Region API then find a mask picture and assign a mask color to be transparent color. Now use the CombineRgn and scan each pixel of the mask picture. If we find the pixel holding the mask color attribute then we have to combine the region which has the same position as the pixel we found and assign the region size as one pixel size ( CreateRegion(x,y,1,1)). After scanning all of pixels of the mask picture ,the non-rectangle region has been done. Now assign the region to the windows to make the window non-rectangle.

But I don't think this a good way to do this. The method is not fast enough. I don't know how to create the region faster. I find VB's UserControl has a powerful way to create a non-rectangle region. Sometime you may create the transparent control with UserControl in VB. You only have to set the backgtound property of UserControl to Transparent and assign a mask picture to the control. After doing this, VB have created the region for you. You can use the SetWindowRgn to change the region of the window. ~Gwyshell


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kwyshell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kwyshell.md)
**Level**          |Unknown
**User Rating**    |4.2 (67 globes from 16 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kwyshell-an-easy-way-to-create-a-transparent-form__1-2296/archive/master.zip)

### API Declarations

```
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
```


### Source Code

```
Now you can test the code in following steps:
 1) Create a new Visual Basic project
 2) Add the UserControl to your project and named it as 'TransparentCtrl'
 3) Add the following code to the control
' Start Control Code
  Public Property Get MaskPicture() As Picture
    Set MaskPicture = UserControl.MaskPicture
  End Property
  Public Property Set MaskPicture(ByVal picNew As Picture)
    Set UserControl.MaskPicture = picNew
    'Put the Refresh() code before the Set Picture Property will
    'have better effection
    Me.Refresh
    Set UserControl.Picture = picNew
    PropertyChanged "MaskPicture"
  End Property
  Public Property Get MaskColor() As OLE_COLOR
    MaskColor = UserControl.MaskColor
  End Property
  Public Property Let MaskColor(ByVal clrMaskColor As OLE_COLOR)
    UserControl.MaskColor = clrMaskColor
    Me.Refresh
    PropertyChanged "MaskColor"
  End Property
  'Refresh() to changed the container region with usercontrol's
  Public Sub Refresh()
    'On Local Error Resume Next
    Dim hRgnNormal As Long
    With UserControl
      If .MaskPicture = 0 Then
        hRgnNormal = CreateRectRgn(0, 0, .ScaleX(.Width), .ScaleY(.Height))
        SetWindowRgn .Extender.Container.hWnd, hRgnNormal, True
      Else
        .Size .ScaleX(.MaskPicture.Width), .ScaleY(.MaskPicture.Height)
        .Extender.Container.Width = .Width
        .Extender.Container.Height = .Height
        .Extender.Move 0, 0
        'Gwyshell
        'Let the system have time to finish the special regions created
        DoEvents
        'Set New Regions
        SetWindowRgn .Extender.Container.hWnd, Me.hRgn , True
        If Err Then
          MsgBox "The Container not support the mothods"
        End If
      End If
    End With
  End Sub
  Public Property Get hRgn() As OLE_HANDLE
    hRgn = CreateRectRgn(0, 0, 1, 1)
    GetWindowRgn Me.hWnd, hRgn
  End Property
  'Following code to persist the control's property
  Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.MaskColor = PropBag.ReadProperty("MaskColor", &H8000000F)
  Set Me.MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
  End Sub
  Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "MaskColor", Me.MaskColor, &H8000000F
    PropBag.WriteProperty "MaskPicture", Me.MaskPicture, Nothing
  End Sub
' End of Control Code
 4) Now close the UserControl Designer to make the control active.
  Add the control on the form and assign the mask picture and mask color
  to the control.
 5) After this, you may see the region of the form has been changed.
 To get the full code please visit here:
http://www.mgt.ncu.edu.tw/~im841150/Documents/TransparentCtrl/TransparentCtrl.htm
```

