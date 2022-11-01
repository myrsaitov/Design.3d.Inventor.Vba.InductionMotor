Option Explicit

Public Sub MoldBaseSample()
   
   'Задаем путь к папке, в которой будут сохраняться файлы деталей
   Dim sFilePath As String
   sFilePath = "D:\WORK\MOLOT\Wiring_VBA\"
   
   'Вызов функций, в которых создаются
   'деталь изготавливаемой отливки (MoldPart.ipt) и
   'базовая деталь полуформы (MoldBase.ipt).
   Call CreateMoldPart(sFilePath & "MoldPart.ipt")
   Call CreateMoldBase(sFilePath & "MoldBase.ipt")
   
   
   'Создание новой детали.
   'Она будет производным компонентом от отливки.
   Dim oPartDoc As PartDocument
   Set oPartDoc = ThisApplication.Documents.Add(kPartDocumentObject, _
   ThisApplication.FileManager.GetTemplateFile(kPartDocumentObject))
      
   ' Создание определяющего объекта (derived definition)
   ' производного компонента.
   Dim oDerivedPartDef As DerivedPartUniformScaleDef
   Set oDerivedPartDef = oPartDoc.ComponentDefinition. _
         ReferenceComponents.DerivedPartComponents. _
         CreateUniformScaleDef(sFilePath & "MoldPart.ipt")
   
   ' Назначение масштабного множителя
   oDerivedPartDef.ScaleFactor = 1.1
   
   
   'В определяющем объекте можно задать и другие
   'опции будущей производной детали.
   'Установки опций по умолчанию нас устраивают во всех
   'случаях, кроме масштабного множителя ScaleFactor.

    
   'Создание производной детали.
   Call oPartDoc.ComponentDefinition.ReferenceComponents. _
            DerivedPartComponents.Add(oDerivedPartDef)
   
   'СОхранить и закрыть производную деталь.
   Call oPartDoc.SaveAs(sFilePath & "ScaledMoldPart.ipt", False)
   
   oPartDoc.Close
   
   
   
   ' Create a new assembly file to put the mold base and scaled part together.
   'Создадим новую сборку. В ней будут собраны вместе деталь и полуформа.
   
   Dim oAsmDoc As AssemblyDocument
   Set oAsmDoc = ThisApplication.Documents.Add(kAssemblyDocumentObject, _
   ThisApplication.FileManager.GetTemplateFile(kAssemblyDocumentObject))
   

   'Создание матрицы для определения положения в сборке ее компонентов.
   'Вновь созданная матрица исходно является единичной.
   'Применение единичной матрицы к вставляемым в сборку компонентам
   'приведет к полному совмещению систем координат деталей и сборки.
      
   Dim oMatrix As Matrix
   Set oMatrix = ThisApplication.TransientGeometry.CreateMatrix
   
   'Позиционирование в сборке полуформы.
   Dim oOcc As ComponentOccurrence
   Set oOcc = oAsmDoc.ComponentDefinition.Occurrences. _
                  Add(sFilePath & "MoldBase.ipt", oMatrix)
   
   'Переименуем компонент, чтобы впоследствии его
   'идентифицировать. Этот метод не является
   'единственным, но он наиболее прост в применении.
   oOcc.Name = "Mold Base"
   
   'Размещение в сборке отмасштабированной детали
   Set oOcc = oAsmDoc.ComponentDefinition.Occurrences. _
                  Add(sFilePath & "ScaledMoldPart.ipt", oMatrix)
   oOcc.Name = "Mold Part"
   
   'Сохраняем и закрываем сборку.
   Call oAsmDoc.SaveAs(sFilePath & "MoldSample.iam", False)
   oAsmDoc.Close
   
   
   'Создаем новую деталь как производную от сборки
   'с вычтьанием тела из полуформы.
   
   'Новая деталь
   Set oPartDoc = ThisApplication.Documents.Add(kPartDocumentObject, _
   ThisApplication.FileManager.GetTemplateFile(kPartDocumentObject))
   
   'Определяющий объект производного компонента - литейной сборки
   Dim oDerivedAsmDef As DerivedAssemblyDefinition
   Set oDerivedAsmDef = oPartDoc.ComponentDefinition. _
         ReferenceComponents.DerivedAssemblyComponents. _
         CreateDefinition(sFilePath & "MoldSample.iam")
   
   'Назначение детали опции вычитания
   oDerivedAsmDef.Occurrences.Item("Mold Part"). _
                     InclusionOption = kDerivedSubtractAll
   
   ' Create the derived assembly.
   'Финал: создание производного компонента от сборки
   Call oPartDoc.ComponentDefinition.ReferenceComponents. _
                     DerivedAssemblyComponents.Add(oDerivedAsmDef)
   
End Sub  'MoldBaseSample
'-------------------------------------------------------------



'Процедура создает деталь, представляющую собой отливку

Private Sub CreateMoldPart(Filename As String)
   
   'Новый документ детали на основе шаблона по умолчанию
   Dim oPartDoc As PartDocument
   Set oPartDoc = ThisApplication.Documents.Add(kPartDocumentObject, _
   ThisApplication.FileManager.GetTemplateFile(kPartDocumentObject))
   
   'Создание эскиза на базовой плоскости XY системы координат
   Dim oSketch As PlanarSketch
   Set oSketch = oPartDoc.ComponentDefinition.Sketches. _
                     Add(oPartDoc.ComponentDefinition.WorkPlanes(3))
   
   Dim oTG As TransientGeometry
   Set oTG = ThisApplication.TransientGeometry
   
   'Геометрия, определяющая форму детали.
   Dim oPoints As ObjectCollection
   Set oPoints = ThisApplication.TransientObjects.CreateObjectCollection
   oPoints.Add oTG.CreatePoint2d(-5, 0)
   oPoints.Add oTG.CreatePoint2d(-4, 3)
   oPoints.Add oTG.CreatePoint2d(-2, 4)
   oPoints.Add oTG.CreatePoint2d(0, 3)
   oPoints.Add oTG.CreatePoint2d(3, 4)
   oPoints.Add oTG.CreatePoint2d(4, 2)
   oPoints.Add oTG.CreatePoint2d(5, 0)
   
   Dim oSpline As SketchSpline
   Set oSpline = oSketch.SketchSplines.Add(oPoints)
   oSpline.FitMethod = kSweetSplineFit
   
   Dim oLine As SketchLine
   Set oLine = oSketch.SketchLines.AddByTwoPoints( _
                  oSpline.FitPoint(1), _
                  oSpline.FitPoint(oSpline.FitPointCount))
   
   Dim oProfile As Profile
   Set oProfile = oSketch.Profiles.AddForSolid
   
   'Создание тела вращения
   Call oPartDoc.ComponentDefinition.Features.RevolveFeatures. _
                     AddFull(oProfile, oLine, kJoinOperation)
   
   'Сохранить и закрыть документ
   Call oPartDoc.SaveAs(Filename, False)
   
   oPartDoc.Close
   
End Sub  'CreateMoldPart
'-------------------------------------------------------------


' Процедура создает заготовку литейной полуформы

Private Sub CreateMoldBase(Filename As String)

   'Новый документ детали на основе шаблона по умолчанию
   Dim oPartDoc As PartDocument
   Set oPartDoc = ThisApplication.Documents.Add(kPartDocumentObject, _
   ThisApplication.FileManager.GetTemplateFile(kPartDocumentObject))
   
   'Создание эскиза на базовой плоскости XY системы координат
   Dim oSketch As PlanarSketch
   Set oSketch = oPartDoc.ComponentDefinition.Sketches. _
                     Add(oPartDoc.ComponentDefinition.WorkPlanes(3))
   
   'Геометрия, определяющая форму детали.
   Dim oTG As TransientGeometry
   Set oTG = ThisApplication.TransientGeometry
   Call oSketch.SketchLines.AddAsTwoPointRectangle( _
                                 oTG.CreatePoint2d(-6, -5), _
                                 oTG.CreatePoint2d(6, 5))
   
   Dim oProfile As Profile
   Set oProfile = oSketch.Profiles.AddForSolid
   
   'Создание тела выдавливанием
   Call oPartDoc.ComponentDefinition.Features. _
            ExtrudeFeatures.AddByDistanceExtent( _
               oProfile, 5, kNegativeExtentDirection, kJoinOperation)
   
   'Сохранить и закрыть документ
   Call oPartDoc.SaveAs(Filename, False)
   
   oPartDoc.Close

End Sub  'CreateMoldBase
'-------------------------------------------------------------