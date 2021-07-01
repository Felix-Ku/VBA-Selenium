Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Dim chrome As Selenium.ChromeDriver

Sub CommandButton1_Click()
    
    Dim Response As Integer
    Dim acac As String
    Dim pwpw As String
    acac = TextBox2.Value
    pwpw = TextBox3.Value
        
    Response = MsgBox("In next step, please select a location with no previous downloaded files, confirm?", vbYesNo)

    If Response = vbYes Then
                'Deal with input text to array section
        
        Dim sFolder As String
        
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = -1 Then ' if OK is pressed
                sFolder = .SelectedItems(1)
            End If
        End With
        
        If sFolder <> "" And acac <> "" And pwpw <> "" Then
        

            MsgBox ("Please wait for the program to automatically download and do not click any buttons in Chrome, close the browser when all downloading is completed.")
        
            Cells(25, 2) = "RUNNING - Location set"
            Cells(27, 2) = sFolder
            
            Dim CntR As Integer
            
            Dim locationIDHead, locationIDTail, locationNameHead, locationNameTail As Integer
            Dim ID, ImageName As String
            
            Dim renamePath, urlPath As String
            
            Dim textValue As String
            
            Dim split1() As String
            Dim subStr As Variant
            
            Dim ids() As String
            Dim names() As String
            
            Dim DQ As String
            
            CntR = 30
            
            DQ = chr(34)
            textValue = Replace(TextBox1.Value, DQ, ".")
            
            split1 = Split(textValue, ",{")
            
            For Each subStr In split1
                
                locationIDHead = InStr(subStr, "///PRIVATE_LINK///")
                locationIDTail = InStr(subStr, "///PRIVATE_LINK///")
                
                locationNameHead = InStr(subStr, "///PRIVATE_LINK///")
                locationNameTail = InStr(subStr, "///PRIVATE_LINK///")
                
                ID = Mid(subStr, locationIDHead + 14, 32)
                ImageName = Mid(subStr, locationNameHead + 17, locationNameTail - locationNameHead - 17)
                
                Cells(CntR, 2) = ID
                Cells(CntR, 3) = ImageName
                
                ReDim Preserve ids(CntR - 30)
                ids(CntR - 30) = ID
                ReDim Preserve names(CntR - 30)
                names(CntR - 30) = ImageName
                
                CntR = CntR + 1
                
                renamePath = sFolder & "\" & ImageName & ".jpg"
                urlPath = "///PRIVATE_LINK///" & ID
                
                
                
            Next subStr
            
            Cells(CntR, 2) = "---END---"
            Cells(CntR, 3) = "---END---"
            
            
            'Auto download section
            
            On Error GoTo Oopss
                Set chrome = New Selenium.ChromeDriver
                
                chrome.SetPreference "download.default_directory", sFolder & "\"
                chrome.SetPreference "download.directory_upgrade", True
                chrome.SetPreference "download.prompt_for_download", False
            
                chrome.Start
                chrome.Get "///PRIVATE_LINK///"
                Sleep 2000
                chrome.FindElementById("///PRIVATE_LINK///").SendKeys acac
                chrome.FindElementById("///PRIVATE_LINK///").SendKeys pwpw
                
                chrome.FindElementByXPath("//span[@class='///PRIVATE_LINK///']").Click
                Sleep 2000
                chrome.FindElementByXPath("//span[@class='///PRIVATE_LINK///']").Click
                Sleep 2000
                chrome.FindElementById("chBoxdeclarationB").Click
                Sleep 2000
                chrome.FindElementByXPath("//span[@class='///PRIVATE_LINK///']").Click
                Sleep 2000
            
            On Error GoTo Oops
                chrome.FindElementByXPath("//span[@class='///PRIVATE_LINK///']").Click
                Debug.Print "Force logout"
Oops:
            
            Debug.Print "Entered"
            
            'Download section
            Dim idss As Variant
            
            For Each idss In ids
                chrome.Get "///PRIVATE_LINK///" & idss
            Next
            
            MsgBox ("Click OK only after you closed the browser.")
                  
            chrome.Quit
            
            
            Dim counterr, i As Integer
            Dim tempChar As String
            Dim namesss As Variant
        

            On Error GoTo Oopss
                MsgBox ("Renaming in progress, please press ok and wait")
                Name sFolder & "\plan.tif" As sFolder & "\" & names(0) & ".tif"
                For i = 1 To UBound(names) - LBound(names)
                    Name sFolder & "\plan " & "(" & i & ").tif" As sFolder & "\" & names(i) & ".tif"
                    Debug.Print i
                Next i
                MsgBox ("Rename completed!!")
            
           
        'End of Auto download section
          
        Else
            Cells(25, 2) = "FAIL - Location not set"
            MsgBox ("FAIL - Location/AC/PW not set!")
        End If
        
        
        
    Else
        MsgBox ("FAIL - Action cancelled!")
    End If
 
Oopss:
    Debug.Print "FAIL - Something went wrong!"
    
End Sub
    