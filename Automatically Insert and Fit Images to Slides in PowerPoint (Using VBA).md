
# 📊 Automatically Insert and Fit Images to Slides in PowerPoint (Using VBA)

This guide shows how to **insert multiple images after a specific slide in PowerPoint** using **VBA (Visual Basic for Applications)** and automatically **resize each image to fit the slide while maintaining its aspect ratio**.

---

## ✅ Use Case

You have a `.pptx` file with 80 slides, and you want to insert a set of images **after slide 57**. Manually doing this would take a lot of time — so we automate it using VBA.

---

## 🛠 Prerequisites

- Microsoft PowerPoint (Desktop version)
- A folder containing your images (e.g., `.jpg`, `.png`)
- Basic familiarity with running a VBA macro

---

## 📁 Folder Setup

Place all your images in a folder.  
For example:

```
C:\Users\riyad\Downloads\Tabulation Method\
```

---

## 🧠 The VBA Macro

1. Open your PowerPoint presentation.
2. Press `Alt + F11` to open the **VBA editor**.
3. Go to `Insert → Module`.
4. Paste the code below:

```vba
Sub InsertPicturesFitToSlide()
    Dim ppt As Presentation
    Dim slideIndex As Integer
    Dim imgPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim newSlide As Slide
    Dim imgShape As Shape
    Dim i As Integer
    Dim slideWidth As Single
    Dim slideHeight As Single

    imgPath = "C:\Users\riyad\Downloads\Tabulation Method\"
    slideIndex = 57

    Set ppt = ActivePresentation
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(imgPath)

    slideWidth = ppt.PageSetup.SlideWidth
    slideHeight = ppt.PageSetup.SlideHeight

    i = 0
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".jpg" Or _
           LCase(Right(file.Name, 4)) = ".png" Or _
           LCase(Right(file.Name, 5)) = ".jpeg" Then

            Set newSlide = ppt.Slides.Add(slideIndex + 1 + i, ppLayoutBlank)
            Set imgShape = newSlide.Shapes.AddPicture(file.Path, _
                MsoTriState.msoFalse, MsoTriState.msoCTrue, 0, 0)

            ' Maintain aspect ratio
            imgShape.LockAspectRatio = msoTrue

            ' Resize based on slide dimensions
            If imgShape.Width / slideWidth > imgShape.Height / slideHeight Then
                ' Wider image — match width
                imgShape.Width = slideWidth
                imgShape.Top = (slideHeight - imgShape.Height) / 2
                imgShape.Left = 0
            Else
                ' Taller image — match height
                imgShape.Height = slideHeight
                imgShape.Left = (slideWidth - imgShape.Width) / 2
                imgShape.Top = 0
            End If

            i = i + 1
        End If
    Next file

    MsgBox i & " images inserted after slide 57 and fit to screen.", vbInformation
End Sub
```

---

## ▶️ How to Run It

1. Press `Alt + F8` in PowerPoint.
2. Select `InsertPicturesFitToSlide`.
3. Click **Run**.
4. Done! 🎉 Your images are now inserted and perfectly fit to the slide.

---

## 🧽 Customization Ideas

- Change `slideIndex = 57` to insert at a different position
- Add captions automatically
- Apply transitions or animations to inserted slides

---

## 💬 Author

Created by [Riyad](https://github.com/) with help from ChatGPT 🤖  
If this helped you, consider starring the repo! ⭐

