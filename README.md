<div align="center">

## Changing Windows Printer w/o using priint dialog


</div>

### Description

How can I change the printer Windows uses in code without using the print common dialog? How can I change orientation?

You can change the printer the VB 3.0 Printer object is pointing to programmatically (without using the common dialogs). Just use the WriteProfileString API call and rewrite the [WINDOWS], DEVICE entry in the WIN.INI file! VB will instantly use the new printer, when the next Printer.Print command is issued. If you get the old printer string before you rewrite it (GetProfileString API call), you can set it back after using a specific printer. This technique is especially useful, when you want to use a FAX printer driver: Select the FAX driver, send your fax by printing to it and switch back to the normal default printer.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB FAQ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-faq.md)
**Level**          |Unknown
**User Rating**    |4.2 (173 globes from 41 users)
**Compatibility**  |
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-faq-changing-windows-printer-w-o-using-priint-dialog__1-78/archive/master.zip)





### Source Code

```
It is recommended (and polite, as we're multitasking) to send a WM_WININCHANGE (&H1A) to all windows to tell them of the change. Also, under some circumstances the printer object won't notice that you have changed the default printer unless you do this.
Declare Function SendMessage(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Global Const WM_WININICHANGE = &H1A
Global Const HWND_BROADCAST = &HFFFF
' Dummy means send to all top windows.
' Send name of changed section as lParam.
lRes = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "Windows")
```

