# 🪑 Root-Cause-Analyses-in-Food-production-facility-using-Excel
This repository contains case studies set in real-life Food production facility that are analyzed using Minitab

---

## 🦺 Root Cause Analysis (RCA)
In Improve phase, a root cause is a factor that causes the defector issue. Removing that factor will prevent the re-occurence of the issue

    Evaluate --> Optimize --> Validate

```mermaid
graph TD;
A[RCA] -->B[Unknown sources of variation];
A -->C[Optimize the process];
A -->D[Implement corrective action];
```

---

## 🗝 Cause and Effect diagram
- It is used to find the root cause and the potential solutions to the problem
- It systematically helps a team break down a problem into several components
- It displays the possible causes in a graphical manner
- It is also known as Fishbone diagram or Ishikawa diagram

### Steps
1. Determine the main classification or headings to group the causes
2. Draw a cause and effect diagram with the problem at the point of the central axis line
3. Brainstorm possible causes of the problem
4. Write the causes on the diagram under classification chosen

### Cause and effect matrix
- C&E matrix helps quantify how much different root causes impact key business effects (like cost, downtime, or defects)
- The prioritization number reflects the effect of each input variable on the output variable
- The process output priority is multiplied with the input variables to arrive at the result for each input variable

### Sample diagram
```mermaid
flowchart TD;
A[Start]-->B[Check Inventory];
B-->C{Inventory Available?};
C-->|Yes| D[Process Order];
C-->|No| E[Request Reorder];
D-->F[Ship Product];
E-->F[Ship Product];
F-->G[End];
```

---

## 📥 The 5 Whys Technique
- Identify the problem and the problem statement
- Arrange for a team brainstorming session
- Explain the purpose
- Analyze the problem and brainstorm backwards
- Ask "Why?" for the answers obtained
- *If a problem occurs, it is usually due to the process and not because of the person or team*

### Corrective action
- A template for capturing customer delivery information
- Eliminate erroneous product deliveries
- Improve quality of service
- Focus on process issues and not people
- Stay within your teams control to change

---

## 👔 Case study: Cause & Effect (Fishbone/Ishikawa) diagram to handle issue in food production facility

### Issue
Frequent production halts in a packaged food line due to packaging defects

### Fishbone categories
- **Man**: Untrained staff, fatigue
- **Machine**: Misaligned sealing machine, low maintenance
- **Material**: Poor packaging quality, wrong bag size
- **Method**: Ineffective SOP, poor inspection process
- **Measurement**: No inline defect tracking, delay in feedback
- **Environment**: High humidity, dust contamination

### VBA macro for Fishbone diagram including sub-causes

    Sub CreateFishboneDiagramWithSubcategories()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Delete all shapes
    ' Dim shp As Object
    ' For Each shp In ws.Shapes
       ' shp.Delete
    ' Next shp
    
    ' Main Spine
    Dim xStart As Single: xStart = 100
    Dim xEnd As Single: xEnd = 700
    Dim yCenter As Single: yCenter = 300
    Dim xStep As Single: xStep = 100
    Dim yOffset As Single: yOffset = 40
    
    ' Add spine
    Dim spine As Shape
    Set spine = ws.Shapes.AddLine(xStart, yCenter, xEnd, yCenter)
    spine.Line.EndArrowheadStyle = msoArrowheadTriangle
    
    ' Categories and subcauses
    Dim categories As Variant
    Dim subcauses As Variant
    
    categories = Array("Man", "Machine", "Method", "Material", "Measurement", "Environment")
    subcauses = Array( _
                Array("Untrained staff", "Fatigue"), _
                Array("Seal misalignment", "Maintenance delay"), _
                Array("No SOP", "Missed checks"), _
                Array("Low quality", "Wrong size"), _
                Array("No inspection", "Delay feedback"), _
                Array("Humidity", "Dust contamination") _
                )
    
    Dim i As Integer, j As Integer
    Dim catX As Single, catY As Single
    Dim branchX As Single, branchY As Single
    Dim angleUp As Boolean
    
    ' Draw branches
    For i = 0 To 5
        angleUp = (i < 3)
        catX = xEnd - xStep * (i + 1)
        catY = yCenter + IIf(angleUp, -1, 1) * yOffset * (i + 1)
        
        ' Main branch
        ws.Shapes.AddLine(catX, catY, xEnd, yCenter).Line.Weight = 1.5
        ws.Shapes.AddTextbox(msoTextOrientationHorizontal, catX - 40, catY + IIf(angleUp, -20, 5), 100, 15).TextFrame.Characters.Text = categories(i)
    
        ' Sub-causes
        For j = 0 To 1
            branchX = catX - 60
            branchY = catY + IIf(angleUp, -1, 1) * 20 * (j + 1)
            
            ' Diagonal line from main branch
            ws.Shapes.AddLine(branchX, branchY, catX, catY).Line.Weight = 1
            ' ws.Shapes.AddTextbox(msoTextOrientationDownward, branchX - 50, branchY + IIf(angleUp, -15, 5), 100, 15).TextFrame.Characters.Text = subcauses(i)(j)
        
        ' Rotated sub-cause text
        Dim txt As Shape
        Set txt = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, branchX - 50, branchY - 10, 80, 20)
        txt.TextFrame.Characters.Text = subcauses(i)(j)
        txt.TextFrame2.Orientation = msoTextOrientationHorizontal
        txt.Rotation = IIf(angleUp, -30, 30)
       Next j
    Next i
    
    ' Add problem statement
    ws.Shapes.AddTextbox(msoTextOrientationHorizontal, xEnd + 10, yCenter - 10, 200, 20).TextFrame.Characters.Text = "Problem: Packaging Defects"
    
    MsgBox "Fishbone Diagram with subcategories created!", vbInformation
    End Sub    

### VBA macro code for CE Matrix

    Sub GenerateCEMatrix()

    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Cells.Clear
    
    ' Setup headers
    ws.Range("A1").Value = "Cause / Effect"
    ws.Range("B1").Value = "Downtime"
    ws.Range("C1").Value = "Defect rate"
    ws.Range("D1").Value = "Cost impact"
    ws.Range("E1").Value = "Total Score"
    
    ' Effect weights
    ws.Range("B2").Value = 5 ' Downtime weight
    ws.Range("C2").Value = 4 ' Defect rate weight
    ws.Range("D2").Value = 3 ' Cost impact weight
    
    ' Example causes
    ws.Range("A3").Value = "Machine misalignment"
    ws.Range("A4").Value = "Training gap"
    ws.Range("A5").Value = "Material Quality"
    ws.Range("A6").Value = "No inspection"
    
    ' Placeholder for scores (user input)
    ' ws.Range("B3:D6").Value = 3 ' default values replace manually
    
    ' Calculate weighted scores and totals
    Dim i As Integer, j As Integer
    For i = 3 To 6 ' Causes
        For j = 2 To 4 ' Effects
            ws.Cells(i, j).Value = 3 ' Default score (1-5 scale)
        Next j
    Next i
        
        ' Sum total score
        ' ws.Cells(i, 5).Formula = "=SUM(B" & i & ":D" & i & ")"
        For i = 3 To 6
            ws.Cells(i, 5).Formula = "=B" & i & "*B2 + C" & i & "*C2+D" & i & "*D2"
    Next i
    
    ' Formatting
    ws.Range("A1:E1").Font.Bold = True
    ws.Columns("A:E").AutoFit
    
    ' Add conditional formatting to total score
    With ws.Range("E3:E6")
            .FormatConditions.AddTop10
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            .FormatConditions(1).TopBottom = x1Top10Top
            .FormatConditions(1).Rank = 1
            .FormatConditions(1).Percent = False
            .FormatConditions(1).Interior.Color = RGB(255, 230, 153) ' Light red
    End With
    
    MsgBox "Cause & Effect matrix created. You can now modify the scores as needed!", vbInformation
    End Sub


            

