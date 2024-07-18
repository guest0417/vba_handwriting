Attribute VB_Name = "handwrite"
Sub handwriting()
    '
    ' handwriting Macro
    '
    
    Dim font_count As Long
    font_count = 6 ' Adjust the number of fonts defined below
    
    Dim font_config() As FontConfig
    ReDim font_config(font_count - 1)
    Dim total_probability As Double
    
    ' Ensure FontConfig class is properly defined
    ' Initialize font configurations
    For i = 0 To (font_count - 1)
        Set font_config(i) = New FontConfig
    Next i
    
    ' Initialize fonts
    font_config(0).InitializeWithValues "美玉体", 16, 0, 0
    font_config(1).InitializeWithValues "方正静蕾简体", 14, 1, 0
    font_config(2).InitializeWithValues "文鼎大钢笔行楷", 14, 1, 1
    font_config(3).InitializeWithValues "汉仪井柏然体简", 17, 1, 0
    font_config(4).InitializeWithValues "华康翩翩体W3P", 15, 1, 0
    font_config(5).InitializeWithValues "BoLeYaYati", 16, 1, 0
    
    ' Calculate total probability
    total_probability = 0
    For i = 0 To (font_count - 1)
        total_probability = total_probability + font_config(i).probability
    Next i
    
    VBA.Randomize
    Dim random As Double
    Dim last_font_ratio As Double
    Dim font_ratio As Double
    Dim font_size As Double
    
    last_font_ratio = 0
    For Each R_Character In ActiveDocument.Characters
        random = VBA.Rnd * total_probability
        
        Dim current_font As FontConfig
        Dim current_count As Double
        current_count = 0
        For i = 0 To (font_count - 1)
            current_count = current_count + font_config(i).probability
            If random < current_count Then
                Set current_font = font_config(i)
                Exit For
            End If
        Next i
    
        font_ratio = last_font_ratio + (VBA.Rnd - 0.5) / 10
         If font_ratio > 0.05 Then
            font_ratio = 0.05
        End If
        If font_ratio < -0.05 Then
            font_ratio = -0.05
        End If
        last_font_ratio = font_ratio
        font_size = current_font.size * (1 + last_font_ratio)
        
        ' Set font properties
        font_size = current_font.size * (1 + last_font_ratio)
        With R_Character.Font
            .name = current_font.name
            .size = font_size
            .Position = (VBA.Rnd - 0.5) * 1.2
            .Spacing = current_font.expanded + VBA.Rnd
        End With
    Next
    
    Application.ScreenUpdating = True
End Sub

