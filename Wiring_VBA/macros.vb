Option Explicit





Public Sub CreateModel()

	Dim sFilePath As String
		sFilePath = "D:\WORK\MOLOT\Wiring_VBA\"
   
	Call CreateStator(sFilePath & "Stator.ipt")
	'Call CreateWinding(sFilePath & "WindingA.ipt")
	
	
	
	
	'Dim oPartDoc As PartDocument
	'Set oPartDoc = ThisApplication.Documents.Add(kPartDocumentObject,ThisApplication.FileManager.GetTemplateFile(kPartDocumentObject))

	'Dim oDerivedPartDef As DerivedPartUniformScaleDef
	'Set oDerivedPartDef = oPartDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateUniformScaleDef(sFilePath & "Stator.ipt")

End Sub

	


Private Sub CreateStator(Filename As String)
	
	Dim STATOR_INNER_D As Double
		STATOR_INNER_D = (180.15)/10 	' cm	D = 180.15 mm

	Dim STATOR_OUTTER_D As Double
		STATOR_OUTTER_D = (355.5)/10	' cm	D = 355.5 mm
	
	Dim STATOR_LENGTH As Double
		STATOR_LENGTH = 20 			' cm
	
	Dim Housing_Radius As Double
		Housing_Radius = (4.8)/10 	'4.8 mm
		
	Dim Housing_Length As Double
		Housing_Length = (47)/10 	'47 mm
	
	Dim STATOR_INNER_OffSet As Double
		STATOR_INNER_OffSet = (2.5)/10 	'2.5 mm
	
	Dim STATOR_HOUSING_COUNT As Integer
		STATOR_HOUSING_COUNT = 48	
	
	Dim STATOR_HOUSING_PHASE As Double  ' First Housing Place Phase in Array
		STATOR_HOUSING_PHASE = 360	
	
	Dim STATOR_COLOR As String
		STATOR_COLOR = "Red"
	
		'STATOR_COLOR = "Yellow"
	
	
	
 
   ' MsgBox "Occurence color: " & oRenderStyle.Name & vbCr 
	
	
	
	Dim oPartDoc As PartDocument
		Set oPartDoc = ThisApplication.Documents.Add(kPartDocumentObject,ThisApplication.FileManager.GetTemplateFile(kPartDocumentObject))
   
	Call CreateStator_Base(oPartDoc,STATOR_INNER_D,STATOR_OUTTER_D,STATOR_LENGTH,STATOR_COLOR)
	
	Call CreateStator_Housing(oPartDoc,STATOR_INNER_D,STATOR_LENGTH,Housing_Radius,Housing_Length,STATOR_INNER_OffSet,STATOR_HOUSING_COUNT,STATOR_COLOR,STATOR_HOUSING_PHASE)

		
	Call oPartDoc.SaveAs(Filename, False)
	oPartDoc.Close
   
End Sub



Private Sub CreateStator_Base(oPartDoc As PartDocument,STATOR_INNER_D As Double,STATOR_OUTTER_D As Double,STATOR_LENGTH As Double,STATOR_COLOR As String)



	Dim oSketch As PlanarSketch
		Set oSketch = oPartDoc.ComponentDefinition.Sketches.Add(oPartDoc.ComponentDefinition.WorkPlanes(3))
   
	Dim oTG As TransientGeometry
		Set oTG = ThisApplication.TransientGeometry

	Dim oCircle1 As SketchCircle
		Set oCircle1 = oSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0,0),STATOR_INNER_D/2)

	Dim oCircle2 As SketchCircle
		Set oCircle2 = oSketch.SketchCircles.AddByCenterRadius(oTG.CreatePoint2d(0,0),STATOR_OUTTER_D/2)
		
	Dim oProfile As Profile
		Set oProfile = oSketch.Profiles.AddForSolid
		
	Dim oExtrude As ExtrudeFeature 
	Set oExtrude = oPartDoc.ComponentDefinition.Features. _
			ExtrudeFeatures.AddByDistanceExtent( _
				oProfile, STATOR_LENGTH, kNegativeExtentDirection, kJoinOperation)
				
				
	Dim oStyle As RenderStyle
	Set oStyle = oPartDoc.RenderStyles.Item(STATOR_COLOR)
	
	Dim oFace As Face
	For Each oFace In oExtrude.Faces
		Call oFace.SetRenderStyle(kOverrideRenderStyle,oStyle)
	Next
	


	
	
	
End Sub




Private Sub CreateStator_Housing(oPartDoc As PartDocument, STATOR_INNER_D As Double,STATOR_LENGTH As Double,R As Double,L As Double,STATOR_INNER_OffSet As Double,STATOR_HOUSING_COUNT As Integer,STATOR_COLOR As String,STATOR_HOUSING_PHASE As Double)

	Dim X_0 As Double
	Dim Y_0 As Double
	
	Dim oSketch As PlanarSketch
	Set oSketch = oPartDoc.ComponentDefinition.Sketches. _
		Add(oPartDoc.ComponentDefinition.WorkPlanes(3))
	
	Dim oTG As TransientGeometry
	Set oTG = ThisApplication.TransientGeometry

	X_0 = STATOR_INNER_D/2 + STATOR_INNER_OffSet 
	Y_0 = 0 

	Dim oArc(1 To 2) As SketchArc
	Set oArc(1) = oSketch.SketchArcs.AddByCenterStartEndPoint( _
		oTG.CreatePoint2d(X_0 + R		,	Y_0			), _
		oTG.CreatePoint2d(X_0 + R		,	Y_0 + R		), _
		oTG.CreatePoint2d(X_0 + R		,	Y_0 - R		))
		
	Dim oLine(1 To 2) As SketchLine
	Set oLine(1) = oSketch.SketchLines.AddByTwoPoints( _
		oArc(1).EndSketchPoint, _
		oTG.CreatePoint2d(X_0 + L - R	,	Y_0 - R		))
   
	Set oArc(2) = oSketch.SketchArcs.AddByCenterStartEndPoint( _
		oTG.CreatePoint2d(X_0 + L - R	,	Y_0			), _
		oLine(1).EndSketchPoint, _
		oTG.CreatePoint2d(X_0 + L - R	,	Y_0 + R		))
				
	Set oLine(2) = oSketch.SketchLines.AddByTwoPoints( _
		oArc(2).EndSketchPoint, _
		oArc(1).StartSketchPoint)
	
	Dim oProfile As Profile
	Set oProfile = oSketch.Profiles.AddForSolid

	Dim oExtrude As ExtrudeFeature
	Set oExtrude = oPartDoc.ComponentDefinition.Features. _
			ExtrudeFeatures.AddByDistanceExtent( _
				oProfile, STATOR_LENGTH, kNegativeExtentDirection, kCutOperation)
				
	Dim oStyle As RenderStyle
	Set oStyle = oPartDoc.RenderStyles.Item(STATOR_COLOR)
	
	Dim oFace As Face
	For Each oFace In oExtrude.Faces
		Call oFace.SetRenderStyle(kOverrideRenderStyle,oStyle)
	Next

	Dim objCol As ObjectCollection
	Set objCol = ThisApplication.TransientObjects.CreateObjectCollection	
		

	'convert degrees to radians
	Dim oDeg As Double
	oDeg = STATOR_HOUSING_PHASE * 0.0174532925  'Start Phase
		
	objCol.Clear
	objCol.Add oExtrude 
	

	Call oPartDoc.ComponentDefinition.Features.CircularPatternFeatures.Add(objCol, oPartDoc.ComponentDefinition.WorkAxes(3), False, STATOR_HOUSING_COUNT, oDeg , True)

	
End Sub







