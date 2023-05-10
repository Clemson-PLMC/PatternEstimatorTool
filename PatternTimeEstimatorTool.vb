'------------------------------------------------------------------------------
' File: PatterTimeEstimatorTool.vb
' Principle Author: John Morris, jhmrrs@clemson.edu, Clemson University
' Initial Creation: 9 May 2023
' This file is owned by the PLM Center at Clemson University. It is made available AS IS 
'   under the attached license. Questions on usage can be directed to plmcenter@clemson.edu
'
' Version History:
' V1 (jhmrrs@clemson.edu, 9 May 2023): Initial version
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'These imports are needed for the following template code
'------------------------------------------------------------------------------
Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.BlockStyler

'------------------------------------------------------------------------------
'Represents Block Styler application class
'------------------------------------------------------------------------------
Public Class PatternTimeEstimator
    'class members
    Private Shared theSession As Session
    Private Shared theUI As UI
    Private Shared theUF as NXOpen.UF.UFSession
    Private theDlxFileName As String
    Private theDialog As NXOpen.BlockStyler.BlockDialog
    Private calibration_label As NXOpen.BlockStyler.Label' Block type: Label
    Private param_group As NXOpen.BlockStyler.Group
    Private select_type As NXOpen.BlockStyler.Enumeration' Block type: Enumeration
    Private num_instances As NXOpen.BlockStyler.IntegerBlock' Block type: Integer
    Private select_group As NXOpen.BlockStyler.Group' Block type: Group
    Private select_feature As NXOpen.BlockStyler.SelectFeature' Block type: Select Feature
    Private select_face As NXOpen.BlockStyler.FaceCollector' Block type: Face Collector
    Private LabelResults As NXOpen.BlockStyler.Label' Block type: Label
    Private data_table As NXOpen.BlockStyler.Table' Block type: TableLayout
    Private Label_FeaturePattern As NXOpen.BlockStyler.Label' Block type: Label
    Private feature_time As NXOpen.BlockStyler.Label' Block type: Label
    Private Label_FacePattern As NXOpen.BlockStyler.Label' Block type: Label
    Private face_time As NXOpen.BlockStyler.Label' Block type: Label
    Private Label_GeoPattern As NXOpen.BlockStyler.Label' Block type: Label
    Private geo_time As NXOpen.BlockStyler.Label' Block type: Label
    Private recalibrate As NXOpen.BlockStyler.Button' Block type: Button
    '------------------------------------------------------------------------------
    'Bit Option for Property: EntityType
    '------------------------------------------------------------------------------
    Public Shared ReadOnly Dim                          EntityType_AllowFaces As Integer =    16
    Public Shared ReadOnly Dim                         EntityType_AllowDatums As Integer =    32
    Public Shared ReadOnly Dim                         EntityType_AllowBodies As Integer =    64
    '------------------------------------------------------------------------------
    'Bit Option for Property: FaceRules
    '------------------------------------------------------------------------------
    Public Shared ReadOnly Dim                           FaceRules_SingleFace As Integer =     1
    Public Shared ReadOnly Dim                          FaceRules_RegionFaces As Integer =     2
    Public Shared ReadOnly Dim                         FaceRules_TangentFaces As Integer =     4
    Public Shared ReadOnly Dim                   FaceRules_TangentRegionFaces As Integer =     8
    Public Shared ReadOnly Dim                            FaceRules_BodyFaces As Integer =    16
    Public Shared ReadOnly Dim                         FaceRules_FeatureFaces As Integer =    32
    Public Shared ReadOnly Dim                        FaceRules_AdjacentFaces As Integer =    64
    Public Shared ReadOnly Dim                  FaceRules_ConnectedBlendFaces As Integer =   128
    Public Shared ReadOnly Dim                        FaceRules_AllBlendFaces As Integer =   256
    Public Shared ReadOnly Dim                             FaceRules_RibFaces As Integer =   512
    Public Shared ReadOnly Dim                            FaceRules_SlotFaces As Integer =  1024
    Public Shared ReadOnly Dim                   FaceRules_BossandPocketFaces As Integer =  2048
    Public Shared ReadOnly Dim                       FaceRules_MergedRibFaces As Integer =  4096
    Public Shared ReadOnly Dim                  FaceRules_RegionBoundaryFaces As Integer =  8192
    Public Shared ReadOnly Dim                 FaceRules_FaceandAdjacentFaces As Integer = 16384
    Public Shared ReadOnly Dim                            FaceRules_HoleFaces As Integer = 32768

    '------------------------------------------------------------------------------
    'Custom Variables
    '------------------------------------------------------------------------------
    Private num_selected_faces      As Integer = 0
    Private num_selected_features   As Integer = 0
    Private feature_scale   As Double = 1
    Private face_scale      As Double = 1
    Private geo_scale       As Double = 1
    Private feature_offset  As Double = 0 
    Private face_offset     As Double = 0 
    Private geo_offset      As Double = 0
    Private high_accuracy_large_sample_mode As Boolean = False   
    
#Region "Block Styler Dialog Designer generator code"
    '------------------------------------------------------------------------------
    'Constructor for NX Styler class
    '------------------------------------------------------------------------------
    Public Sub New()
        Try
        
            theSession = Session.GetSession()
            theUI = UI.GetUI()
            theUF = NXOpen.UF.UFSession.GetUFSession()
            theDlxFileName = "PatternTimeEstimator.dlx"
            theDialog = theUI.CreateDialog(theDlxFileName)
            theDialog.AddUpdateHandler(AddressOf update_cb)
            theDialog.AddInitializeHandler(AddressOf initialize_cb)
            theDialog.AddDialogShownHandler(AddressOf dialogShown_cb)
        
        Catch ex As Exception
        
            '---- Enter your exception handling code here -----
            Throw ex
        End Try
    End Sub
#End Region
    
    '------------------------------- DIALOG LAUNCHING ---------------------------------
    '
    '    Before invoking this application one needs to open any part/empty part in NX
    '    because of the behavior of the blocks.
    '
    '    Make sure the dlx file is in one of the following locations:
    '        1.) From where NX session is launched
    '        2.) $UGII_USER_DIR/application
    '        3.) For released applications, using UGII_CUSTOM_DIRECTORY_FILE is highly
    '            recommended. This variable is set to a full directory path to a file 
    '            containing a list of root directories for all custom applications.
    '            e.g., UGII_CUSTOM_DIRECTORY_FILE=$UGII_BASE_DIR\ugii\menus\custom_dirs.dat
    '
    '    You can create the dialog using one of the following way:
    '
    '    1. Journal Replay
    '
    '        1) Replay this file through Tool->Journal->Play Menu.
    '
    '    2. USER EXIT
    '
    '        1) Create the Shared Library -- Refer "Block UI Styler programmer's guide"
    '        2) Invoke the Shared Library through File->Execute->NX Open menu.
    '
    '------------------------------------------------------------------------------
    Public Shared Sub Main()
        Dim thePatternTimeEstimator As PatternTimeEstimator = Nothing
        Try
        
            thePatternTimeEstimator = New PatternTimeEstimator()
            ' The following method shows the dialog immediately
            thePatternTimeEstimator.Show()
        
        Catch ex As Exception
        
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        Finally
            If thePatternTimeEstimator IsNot Nothing Then 
                thePatternTimeEstimator.Dispose()
                thePatternTimeEstimator = Nothing
            End If
        End Try
    End Sub
    '------------------------------------------------------------------------------
    ' This method specifies how a shared image is unloaded from memory
    ' within NX. This method gives you the capability to unload an
    ' internal NX Open application or user  exit from NX. Specify any
    ' one of the three constants as a return value to determine the type
    ' of unload to perform:
    '
    '
    '    Immediately : unload the library as soon as the automation program has completed
    '    Explicitly  : unload the library from the "Unload Shared Image" dialog
    '    AtTermination : unload the library when the NX session terminates
    '
    '
    ' NOTE:  A program which associates NX Open applications with the menubar
    ' MUST NOT use this option since it will UNLOAD your NX Open application image
    ' from the menubar.
    '------------------------------------------------------------------------------
    Public Shared Function GetUnloadOption(ByVal arg As String) As Integer
        'Return CType(Session.LibraryUnloadOption.Explicitly, Integer)
         Return CType(Session.LibraryUnloadOption.Immediately, Integer)
        ' Return CType(Session.LibraryUnloadOption.AtTermination, Integer)
    End Function
    '------------------------------------------------------------------------------
    ' Following method cleanup any housekeeping chores that may be needed.
    ' This method is automatically called by NX.
    '------------------------------------------------------------------------------
    Public Shared Sub UnloadLibrary(ByVal arg As String)
        Try
        
        
        Catch ex As Exception
        
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub
    
    '------------------------------------------------------------------------------
    'This method shows the dialog on the screen
    '------------------------------------------------------------------------------
    Public Sub Show()
        Try
        
            theDialog.Show
        
        Catch ex As Exception
        
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub
    
    '------------------------------------------------------------------------------
    'Method Name: Dispose
    '------------------------------------------------------------------------------
    Public Sub Dispose()
        If theDialog IsNot Nothing Then 
            theDialog.Dispose()
            theDialog = Nothing
        End If
    End Sub
    
    '------------------------------------------------------------------------------
    '---------------------Block UI Styler Callback Functions--------------------------
    '------------------------------------------------------------------------------
    
    '------------------------------------------------------------------------------
    'Callback Name: initialize_cb
    '------------------------------------------------------------------------------
    Public Sub initialize_cb()
        Try
            calibration_label = CType(theDialog.TopBlock.FindBlock("calibration_label"), NXOpen.BlockStyler.Label)
            param_group = CType(theDialog.TopBlock.FindBlock("param_group"), NXOpen.BlockStyler.Group)
            select_type = CType(theDialog.TopBlock.FindBlock("select_type"), NXOpen.BlockStyler.Enumeration)
            num_instances = CType(theDialog.TopBlock.FindBlock("num_instances"), NXOpen.BlockStyler.IntegerBlock)
            select_group = CType(theDialog.TopBlock.FindBlock("select_group"), NXOpen.BlockStyler.Group)
            select_feature = CType(theDialog.TopBlock.FindBlock("select_feature"), NXOpen.BlockStyler.SelectFeature)
            select_face = CType(theDialog.TopBlock.FindBlock("select_face"), NXOpen.BlockStyler.FaceCollector)
            LabelResults = CType(theDialog.TopBlock.FindBlock("LabelResults"), NXOpen.BlockStyler.Label)
            data_table = CType(theDialog.TopBlock.FindBlock("data_table"), NXOpen.BlockStyler.Table)
            Label_FeaturePattern = CType(theDialog.TopBlock.FindBlock("Label_FeaturePattern"), NXOpen.BlockStyler.Label)
            feature_time = CType(theDialog.TopBlock.FindBlock("feature_time"), NXOpen.BlockStyler.Label)
            Label_FacePattern = CType(theDialog.TopBlock.FindBlock("Label_FacePattern"), NXOpen.BlockStyler.Label)
            face_time = CType(theDialog.TopBlock.FindBlock("face_time"), NXOpen.BlockStyler.Label)
            Label_GeoPattern = CType(theDialog.TopBlock.FindBlock("Label_GeoPattern"), NXOpen.BlockStyler.Label)
            geo_time = CType(theDialog.TopBlock.FindBlock("geo_time"), NXOpen.BlockStyler.Label)
            recalibrate = CType(theDialog.TopBlock.FindBlock("recalibrate"), NXOpen.BlockStyler.Button)

        Catch ex As Exception
        
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub
    
    '------------------------------------------------------------------------------
    'Callback Name: dialogShown_cb
    'This callback is executed just before the dialog launch. Thus any value set 
    'here will take precedence and dialog will be launched showing that value. 
    '------------------------------------------------------------------------------
    Public Sub dialogShown_cb()
        Try
            ' Set up model
            calibrate
            updateFaceRules
        
        Catch ex As Exception
        
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub
    
    '------------------------------------------------------------------------------
    'Callback Name: update_cb
    '------------------------------------------------------------------------------
    Public Function update_cb(ByVal block As NXOpen.BlockStyler.UIBlock) As Integer
        Try
            If block Is calibration_label Then
        
            ElseIf block Is select_type Then
            updateFaceRules
            
            ElseIf block Is num_instances Then
            updateTimeLabels

            ElseIf block Is select_feature Then
            getObjectCount
            updateTimeLabels
            
            ElseIf block Is select_face Then
            getObjectCount
            updateTimeLabels
            
            ElseIf block Is LabelResults Then
            '---- Enter your code here -----
            
            ElseIf block Is Label_FeaturePattern Then
            '---- Enter your code here -----
            
            ElseIf block Is feature_time Then
            '---- Enter your code here -----
            
            ElseIf block Is Label_FacePattern Then
            '---- Enter your code here -----
            
            ElseIf block Is face_time Then
            '---- Enter your code here -----
            
            ElseIf block Is Label_GeoPattern Then
            '---- Enter your code here -----
            
            ElseIf block Is geo_time Then
            '---- Enter your code here -----

            ElseIf block Is recalibrate Then
                Dim recal_msg As String, proceed As Integer = 2
                recal_msg = "Recalibration is more accurate for large patterns (that might cause CPU throttling) but can take a long time. Continue anyway?"
                proceed = theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Question, recal_msg)
                If proceed = 1 Then
                    high_accuracy_large_sample_mode = True
                    calibrate
                    getObjectCount
                    updateTimeLabels
                Else
                    high_accuracy_large_sample_mode = False
                End If

            End If
        
        Catch ex As Exception
        
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
        update_cb = 0
    End Function
    
    '------------------------------------------------------------------------------
    'Function Name: GetBlockProperties
    'Returns the propertylist of the specified BlockID
    '------------------------------------------------------------------------------
    Public Function GetBlockProperties(ByVal blockID As String) As PropertyList
        GetBlockProperties = Nothing
        Try
        
            GetBlockProperties = theDialog.GetBlockProperties(blockID)
        
        Catch ex As Exception
        
            '---- Enter your exception handling code here -----
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Function

    '------------------------------------------------------------------------------
    'Function Name: RemoveCalibMsg
    'Removes the calibration message and prepares the Block UI for primary usage
    '------------------------------------------------------------------------------
    Private Sub RemoveCalibMsg()
        calibration_label.Show = False
        param_group.Enable = True
        select_group.Enable = True
        data_table.Enable = True
    End Sub

    '------------------------------------------------------------------------------
    'Function Name: ShowCalibMsg
    'Sets the calibration message and locks out the Block UI
    '------------------------------------------------------------------------------
    Private Sub ShowCalibMsg()
        calibration_label.Show = True
        param_group.Enable = False
        select_group.Enable = False
        data_table.Enable = False
    End Sub

    '------------------------------------------------------------------------------
    'Function Name: calibrate
    'Finds scaling factors for each patterning type
    '------------------------------------------------------------------------------
    Private Sub calibrate(Optional debug_mode As Boolean = False)
            ShowCalibMsg
            Dim markId1 As NXOpen.Session.UndoMarkId = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Delete")
            Dim SPACING as Decimal = 2.5
            Dim block As NXOpen.Features.Block = Nothing
            Dim slab as NXOpen.Features.Block = Nothing
            createInitialObjects(block, slab, debug_mode)

            runTests("FEATURE", block, SPACING, feature_scale, feature_offset, debug_mode)
            runTests("FACE", block, SPACING, face_scale, face_offset, debug_mode)
            If Not debug_mode Then
                theSession.UpdateManager.AddToDeleteList(slab)
            End If
            runTests("GEO", block, SPACING, geo_scale, geo_offset, debug_mode)

            If Not debug_mode Then
                theSession.UpdateManager.AddToDeleteList(block)
                theSession.UpdateManager.DoUpdate(markId1)
            End If

            RemoveCalibMsg
    End Sub

    '------------------------------------------------------------------------------
    'Function Name: createInitialObjects
    'Creates the block that is used as the seed feature (or geometry) for all the patterns. Also
    ' creates a large block (named slab) used for joining features together.
    '------------------------------------------------------------------------------
    Private Sub createInitialObjects(ByRef block As NXOpen.Features.Block, ByRef slab As NXOpen.Features.Block, Optional debug_mode As Boolean = False)
        Dim remoteness_scale As Decimal = 10000
        If debug_mode Then 
            remoteness_scale = 1 
        End If
        Dim location As Decimal = -200 * remoteness_scale
        Dim by_location As Decimal = location - 0.2
        Dim remote_point as NXOpen.Point3d = New NXOpen.Point3d(location, location, location)
        Dim by_remote_point as NXOpen.Point3d = New NXOpen.Point3d(by_location, by_location, by_location)

        block = makeBlock(remote_point)
        slab = makeSlab(by_remote_point, block)
    End Sub

    '------------------------------------------------------------------------------
    'Function Name: makeBlock
    'Makes a single cube
    '------------------------------------------------------------------------------
    Private Function makeBlock(center as NXOpen.Point3d, Optional length As String = "1") As NXOpen.Features.Block
        Try
            Dim workPart As NXOpen.Part = theSession.Parts.Work
            Dim builder As NXOpen.Features.BlockFeatureBuilder
            builder = workPart.Features.CreateBlockFeatureBuilder(Nothing)

            builder.SetOriginAndLengths(center, length, length, length)
            builder.BooleanOption.Type = NXOpen.GeometricUtilities.BooleanOperation.BooleanType.Create
            Dim block As NXOpen.Features.Block = builder.CommitFeature
            
            builder.Destroy
            Return block

        Catch ex As Exception
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Function

    '------------------------------------------------------------------------------
    'Function Name: makeSlab
    'Makes a large slab to do patterning on
    '------------------------------------------------------------------------------
    Private Function makeSlab(center as NXOpen.Point3d, block as NXOpen.Features.Block, Optional length as String = "1500", Optional thickness as String = ".2") As NXOpen.Features.Block
        Try
            Dim workPart As NXOpen.Part = theSession.Parts.Work
            Dim slab_builder As NXOpen.Features.BlockFeatureBuilder
            slab_builder = workPart.Features.CreateBlockFeatureBuilder(Nothing)

            slab_builder.SetOriginAndLengths(center, length, length, thickness)
            Dim block_body_feature As NXOpen.Features.BodyFeature = CType(block, NXOpen.Features.BodyFeature)
            Dim block_bodies() As NXOpen.Body = block_body_feature.GetBodies()
            slab_builder.setBooleanOperationAndTarget(NXOpen.GeometricUtilities.BooleanOperation.BooleanType.Unite, block_bodies(0))
            
            Dim slab As NXOpen.Features.Block = slab_builder.CommitFeature
            slab_builder.Destroy
            Return slab

        Catch ex As Exception
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Function

    '------------------------------------------------------------------------------
    'Function Name: runTests
    'Runs the calibration test for the type of pattern passed as a String and sets the associated constants.
    ' Type is an enumerated variable with values:
    ' - "FEATURE": for a feature pattern
    ' - "FACE": for a face pattern
    ' - "GEO": for a Geo pattern
    '------------------------------------------------------------------------------
    Private Sub runTests(type as String, block As NXOpen.Features.Block, spacing As Decimal, ByRef scale As Double, ByRef offset As Double, Optional debug_mode As Boolean = False)
        Dim duration1 as Single, duration2 as Single
        Dim NUM_TEST1_INST As Integer = 5, NUM_TEST2_INST As Integer = 10
        If high_accuracy_large_sample_mode Then
            Dim sample_base As Integer = 1200
            If num_instances.Value >= 400 And num_instances.Value <= 1200 Then
                sample_base = num_instances.Value
            End If
            NUM_TEST1_INST = sample_base - Int(sample_base * 0.1)
            NUM_TEST2_INST= sample_base + Int(sample_base * 0.1)
        End If
        Dim NUM_OBJECTS_IN_CUBE As Integer = 5 '6 faces in test cube minus bottom face merged with slab
        
        runTest(type, block, 2, spacing, False) 'Eliminate feature startup bias
        duration1 = runTest(type, block, NUM_TEST1_INST, spacing, debug_mode)
        duration2 = runTest(type, block, NUM_TEST2_INST, spacing, debug_mode)

        scale = (duration2 - duration1) / (NUM_OBJECTS_IN_CUBE * (NUM_TEST2_INST - NUM_TEST1_INST))
        offset = duration1 - (scale * NUM_TEST1_INST * NUM_OBJECTS_IN_CUBE)
    End Sub

    '------------------------------------------------------------------------------
    'Function Name: runTest
    'Calls the correct calibration function for the type of pattern passed as a String. Returns the duration of the test
    ' Type is an enumerated variable with values:
    ' - "FEATURE": for a feature pattern
    ' - "FACE": for a face pattern
    ' - "GEO": for a Geo pattern
    '------------------------------------------------------------------------------
    Private Function runTest(type as String, block As NxOpen.Features.Block, num_instances As Integer, spacing As Decimal, Optional debug_mode As Boolean = False) as Single
        Dim start_time as Single, duration as Single

        start_time = Timer
        If type = "FEATURE" Then
            Dim feature_pattern As NXOpen.Features.PatternFeature = makeFeaturePattern(block, num_instances, spacing)
            duration = Timer - start_time
            theSession.UpdateManager.AddToDeleteList(feature_pattern)
        ElseIf type = "FACE" Then
            Dim face_pattern As NXOpen.Features.PatternFaceFeature = makeFacePattern(block, num_instances, spacing)
            duration = Timer - start_time
            theSession.UpdateManager.AddToDeleteList(face_pattern)
        ElseIf type = "GEO" Then 
            Dim geo_pattern As NXOpen.Features.PatternGeometry = makeGeoPattern(block, num_instances, spacing)
            duration = Timer - start_time
            theSession.UpdateManager.AddToDeleteList(geo_pattern)
        Else 
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, "Incorrect input (" & type & ") to runTest function")
        End If

        If debug_mode Then
            Guide.InfoWriteLine(type & " Pattern elapsed time (" & num_instances & "): " & CStr(duration))
        End If

        return duration
    End Function

    '------------------------------------------------------------------------------
    'Function Name: makeFeaturePattern
    'Makes Feature Pattern from block
    '------------------------------------------------------------------------------
    Private Function makeFeaturePattern(block As NXOpen.Features.Block, num_instances as String, spacing as String) As NXOpen.Features.PatternFeature
        Try
            Dim workPart As NXOpen.Part = theSession.Parts.Work
            Dim builder As NXOpen.Features.PatternFeatureBuilder

            builder = workPart.Features.CreatePatternFeatureBuilder(Nothing)
            builder.FeatureList.Add(block)
            builder.PatternService.RectangularDefinition.XSpacing.NCopies.SetFormula(num_instances)
            builder.PatternService.RectangularDefinition.XSpacing.PitchDistance.SetFormula(spacing)
            builder.PatternService.RectangularDefinition.XDirection = makeDirection
            
            return builder.Commit()

        Catch ex As Exception
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Function

    '------------------------------------------------------------------------------
    'Function Name: makeFacePattern
    'Makes Face Pattern from block
    '------------------------------------------------------------------------------
    Private Function makeFacePattern(block As NXOpen.Features.Block, num_instances as String, spacing as String) As NXOpen.Features.PatternFaceFeature
        Try
            Dim workPart As NXOpen.Part = theSession.Parts.Work
            Dim builder As NXOpen.Features.PatternFaceFeatureBuilder
            builder = workPart.Features.CreatePatternFaceFeatureBuilder(Nothing)

            Dim faceFeatureRule As NXOpen.FaceFeatureRule = workPart.ScRuleFactory.CreateRuleFaceFeature({block})
            builder.FaceCollector.ReplaceRules({faceFeatureRule}, False)

            builder.PatternDefinition.RectangularDefinition.XDirection = makeDirection
            builder.PatternDefinition.RectangularDefinition.XSpacing.NCopies.SetFormula(num_instances)
            builder.PatternDefinition.RectangularDefinition.XSpacing.PitchDistance.SetFormula(spacing)
                        
            return builder.Commit()

        Catch ex As Exception
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Function

    '------------------------------------------------------------------------------
    'Function Name: makeGeoPattern
    'Makes a Geometry Pattern from block
    '------------------------------------------------------------------------------
    Private Function makeGeoPattern(block As NXOpen.Features.Block, num_instances as String, spacing as String) As NXOpen.Features.PatternGeometry
        Try
            Dim workPart As NXOpen.Part = theSession.Parts.Work
            Dim builder As NXOpen.Features.PatternGeometryBuilder
            builder = workPart.Features.CreatePatternGeometryBuilder(Nothing)

            Dim scCollector As NXOpen.ScCollector = workPart.ScCollectors.CreateCollector()
            Dim faceDumbRule as NXOpen.FaceDumbRule = workPart.ScRuleFactory.CreateRuleFaceDumb(block.getFaces())
            scCollector.ReplaceRules({faceDumbRule}, False)
            builder.GeometryToPattern.Add(scCollector)

            builder.PatternService.RectangularDefinition.XDirection = makeDirection
            builder.PatternService.RectangularDefinition.XSpacing.NCopies.SetFormula(num_instances)
            builder.PatternService.RectangularDefinition.XSpacing.PitchDistance.SetFormula(spacing)

            Return builder.Commit()

        Catch ex As Exception
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Function

    '------------------------------------------------------------------------------
    'Function Name: makeDirection
    'Makes and returns a direction from a vector
    '------------------------------------------------------------------------------
    Private Function makeDirection() As NXOpen.Direction
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim origin As New Point3d(0.0, 0.0, 0.0)
        Dim vector As New Vector3d(1.0, 0.0, 0.0)
        return workPart.Directions.CreateDirection(origin, vector, NXOpen.SmartObject.UpdateOption.WithinModeling)
    End Function
    
    '------------------------------------------------------------------------------
    'Function Name: updateFaceRules
    'Updates the face rules based on selection criteria
    '------------------------------------------------------------------------------
    Private Sub updateFaceRules() 
        Try
            If select_type.ValueAsString = "Feature(s)" Then
                select_feature.Show = True
                select_face.Show = False
                select_face.FaceRules = FaceRules_FeatureFaces
                select_face.LabelString = "Select Feature Faces"
            Else
                select_feature.Show = False
                select_face.Show = True
                select_face.FaceRules = FaceRules_SingleFace
                select_face.LabelString = "Select Faces"
            End If

        Catch ex As Exception
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    '------------------------------------------------------------------------------
    'Function Name: getObjectCount
    'Counts the number of faces and edges in the selection
    '----------------------------------------------------------------------------
    Private Sub getObjectCount()
        Try
            Dim features() As NXOpen.TaggedObject, this_feat as NXOpen.Features.Feature, total_faces As Integer = 0
            features = select_feature.GetSelectedObjects()
            num_selected_features = features.Length

            Dim feat_faces() As Tag, feature_tag As Tag
            For Each this_feat in features
                feature_tag = this_feat.Tag
                theUF.modl.AskFeatFaces(feature_tag, feat_faces)
                total_faces = total_faces + feat_faces.Length
                Next this_feat

            Dim faces() As NXOpen.TaggedObject, this_face as NXOpen.Face
            faces = select_face.GetSelectedObjects()
            total_faces = total_faces + UBound(faces) - LBound(faces) + 1

            ' Dim total_edges As Integer = 0, edges() as NXOpen.Edge, num_edges As Integer
            ' For Each this_face in faces
            '     edges = this_face.getEdges()
            '     num_edges = UBound(edges) - LBound(edges) + 1
            '     total_edges += num_edges
            '     Next this_face

            num_selected_faces = total_faces

        Catch ex As Exception
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    '------------------------------------------------------------------------------
    'Function Name: updateTimeLabels
    'Updates labels with the estimated time to complete each pattern
    '------------------------------------------------------------------------------
    Private Sub updateTimeLabels()
        Try
            If num_selected_faces = 0 Then
                feature_time.Label = "None Selected"
                face_time.Label = "None Selected "
                geo_time.Label = "None Selected"
            Else
                feature_time.Label = formatTime(feature_scale * num_selected_faces * num_instances.Value + feature_offset)
                face_time.Label = formatTime(face_scale * num_selected_faces * num_instances.Value + face_offset)
                geo_time.Label = formatTime(geo_scale * num_selected_faces * num_instances.Value + geo_offset)
            End If

        Catch ex As Exception
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString)
        End Try
    End Sub

    '------------------------------------------------------------------------------
    'Function Name: formatTime
    'Formats the time with a uniform standard
    '------------------------------------------------------------------------------
    Private Function formatTime(time as Decimal) as String
        Dim out As String
        Dim hours As Integer = Int(time / 3600)
        Dim minutes As Integer = Int((time Mod 3600) / 60)
        Dim seconds As Decimal = time Mod 60

        If time < 10
            return "<10s"
        ElseIf Not hours = 0 Then
            return hours & "h " & minutes & "m "
        ElseIf Not minutes = 0 Then
            return minutes & "m " & CINT(seconds) & "s"
        Else
            return CINT(seconds) & "s"
        End If
    End Function

End Class
