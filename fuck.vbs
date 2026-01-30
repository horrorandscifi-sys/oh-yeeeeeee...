' oh_yeeeeeee.vbs
Option Explicit

Dim fso, file, article

article = "=============================" & vbCrLf & _
          "      Oh yeeeeeee..." & vbCrLf & _
          "=============================" & vbCrLf & vbCrLf & _
          "The Oh yeeeeeee... collection is a unique archive of files from different operating systems." & vbCrLf & _
          "It brings together resources and materials from:" & vbCrLf & _
          "- MacOS" & vbCrLf & _
          "- Windows 10–11" & vbCrLf & _
          "- Linux" & vbCrLf & _
          "- iOS" & vbCrLf & _
          "- Android" & vbCrLf & _
          "- Chrome OS" & vbCrLf & vbCrLf & _
          "This collection serves as a cross‑platform library, allowing enthusiasts and developers to explore," & vbCrLf & _
          "compare, and experiment with files originating from diverse system environments." & vbCrLf & vbCrLf & _
          "Whether you are a researcher, a hobbyist, or simply curious about how different operating systems" & vbCrLf & _
          "structure their files, the Oh yeeeeeee... collection provides a fascinating window into the digital" & vbCrLf & _
          "ecosystems we use every day." & vbCrLf

' Выводим статью в окно сообщения
MsgBox article, vbOKOnly, "Oh yeeeeeee... Article"

' Сохраняем статью в файл
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile("OhYeeeeeee_Article.txt", True)
file.Write article
file.Close

MsgBox "Статья сохранена в файл OhYeeeeeee_Article.txt", vbInformation, "Готово"
