# eVB-Notification-Icon
This module enables your eVB application to add an icon to the taskbar status area (or systray). You can respond to taps and double-taps, currently right 'click' doesn't seem to work.

While eVB does not have the ability to subclass forms and inspect or modify window messages, the way this works is by (ab)using the `MouseDown` event of the form. The [Shell_NotifyIcon](https://learn.microsoft.com/en-us/previous-versions/ms942613(v=msdn.10)) API function lets you specify which window message to send when the user interacts with your icon. eVB is already handling the mouse related event messages so this message will appear as a regular `MouseDown` message to your form. The example project includes code to differentiate between regular `MouseDown` messages and ones triggered from the notify icon, so you'll be able to handle both appropriately.

This method is similar to the approach Microsoft used in [Q176085](https://jeffpar.github.io/kbarchive/kb/176/Q176085/), which I used as a reference for this module. However, unlike the article, I had to use the `MouseDown` event instead of `MouseMove`, because Windows CE seems to handle these messages differently and sending `WM_MOUSEMOVE` does not trigger the corresponding event in eVB.

_An interesting side note is that the Microsoft article claims that this functionality is possible due to the new abilities of VB5/VB6 (most notably the AddressOf operator). However, the given example does not make use of these functionalities at all, and in theory would work just as well in VB4 (or indeed, eVB)._

To use this module, add an ImageList to your form, add the image(s) as per normal, then call the `NotifyIcon_Add` function whenever you want to show the icon.  You can use `NotifyIcon_Modify` to change the currently displayed icon, and `NotifyIcon_Remove` to remove it from the tray.  The example project demonstrates all of these functions.

Make sure to always call `NotifyIcon_Remove` when your application ends, or the icon will remain until the user taps on it to make it go away, in addition, memory may be leaked because icon handles will not be properly cleaned up.

Currently, only one icon is supported per application.

A complete application could look like this:

```vb
Option Explicit

Private Sub Form_Load()

    ImageList.Add "icon_small.bmp"

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'This procedure receives the callbacks from the System Tray icon.
    Dim lngResult  As Long

    Dim lngMessage As Long

    'The value of X will vary depending upon the scalemode setting.
    If ScaleMode = vbPixels Then
        lngMessage = X
    Else
        lngMessage = X / Screen.TwipsPerPixelX
    End If

    Select Case lngMessage

        Case WM_LBUTTONUP 'The user has tapped on the icon once.
            Show
            NotifyIcon_Remove

        Case WM_LBUTTONDBLCLK 'The user has double-tapped on the icon.
            Show
            NotifyIcon_Remove

        Case WM_RBUTTONUP 'The user has tapped on the icon while holding Ctrl, does not seem to work for now.
            Show
            NotifyIcon_Remove
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Hide
    NotifyIcon_Add hWnd, ImageList.hImageList, 0
    Cancel = 1
End Sub
```

## Screenshots

![Screenshot showing the example application, the notification area icon is currently not visible.](https://github.com/WinCEDev/eVB-Notification-Icon/blob/main/Screenshots/CAPT0000.png?raw=1)

![Screenshot showing the example application, the notification area icon is currently visible.](https://github.com/WinCEDev/eVB-Notification-Icon/blob/main/Screenshots/CAPT0001.png?raw=1)

## Links

- [HPC:Factor Forum Thread](https://www.hpcfactor.com/forums/forums/thread-view.asp?tid=20861&posts=1)
