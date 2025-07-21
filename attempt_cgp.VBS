'===============================================================================
'  8‑POLE MULTI‑LAYER STRIPLINE FILTER  –  NO VIAS, MIRRORED LAYOUT
'  Author:  …                                                                   
'  Last rev: 2025‑07‑21                                                         
'===============================================================================

'------------------------------------------------------------------------------
' 0.  HFSS OBJECTS
'------------------------------------------------------------------------------
Dim oApp, oDesktop, oProject, oDesign, oEditor, oDefs, oSetup
Set oApp      = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop  = oApp.GetAppDesktop()
oDesktop.RestoreWindow
Set oProject  = oDesktop.NewProject
Set oProject  = oDesktop.GetActiveProject
oProject.InsertDesign "HFSS", "Stripline8Pole", "DrivenModal", ""
Set oDesign   = oProject.GetActiveDesign
Set oEditor   = oDesign.SetActiveEditor("3D Modeler")
Set oDefs     = oProject.GetDefinitionManager()

'------------------------------------------------------------------------------
' 1.  UNITS
'------------------------------------------------------------------------------
oEditor.SetModelUnits Array("NAME:Units Parameter","Units:=","mil","Rescale:=",false)

'------------------------------------------------------------------------------
' 2.  MATERIAL DEFINITIONS  (lossy)
'------------------------------------------------------------------------------
oDefs.AddMaterial Array( _
  "NAME:Rogers3010"         , "CoordinateSystemType:=","Cartesian",Array("NAME:AttachedData"), _
  "permittivity:="          , "10.2", _
  "dielectric_loss_tangent:=", "0.0023" )

oDefs.AddMaterial Array( _
  "NAME:Rogers2929"         , "CoordinateSystemType:=","Cartesian",Array("NAME:AttachedData"), _
  "permittivity:="          , "2.94", _
  "dielectric_loss_tangent:=", "0.002" )

oDefs.AddMaterial Array( _
  "NAME:Copper_HP"          , "CoordinateSystemType:=","Cartesian",Array("NAME:AttachedData"), _
  Array("NAME:ModifierData"), "conductivity:=", "5.8e7" )

'------------------------------------------------------------------------------
' 3.  GLOBAL DESIGN VARIABLES  (edit these in HFSS > Properties > Variables)
'------------------------------------------------------------------------------
Dim vars: vars = Array( _   'NAME ,                value
  Array("fc_GHz"               , "12.0" ), _                   ' centre freq
  Array("StartFreq_GHz"        , " 8.0" ), _
  Array("StopFreq_GHz"         , "16.0" ), _

  ' — stack‑up —
  Array("Core_H1"              , "8"   ), _                    ' mil  (top core)
  Array("Prepreg_H"            , "4"   ), _
  Array("Core_H2"              , "8"   ), _                    ' mil  (bottom core)
  Array("Cond_Top"             , "0.7" ), _                    ' mil
  Array("Cond_Trace"           , "0.7" ), _
  Array("Cond_Bot"             , "0.7" ), _

  ' — resonator dims (nominal) —
  Array("ResA_side"            , "140" ), _
  Array("ResA_slot_W"          , "20"  ), _
  Array("ResA_slot_L"          , "60"  ), _
  Array("ResB_W"               , "160" ), _
  Array("ResB_H"               , "120" ), _
  Array("ResB_slot_W"          , "20"  ), _
  Array("ResB_slot_sep"        , "30"  ), _                    ' gap between dual slots
  Array("Trace_W"              , "10"  ), _                    ' all resonators

  ' — spacing  —
  Array("Res_x_spacing"        , "60"  ), _
  Array("Res_y_spacing"        , "80"  ), _
  Array("XYSpacing_scale"      , "1"   ), _                    ' global squeeze / stretch
  Array("Z_spacing"            , "Core_H1 + Prepreg_H/2" ), _  ' used for inter‑layer

  ' — launch / ports —
  Array("Port_Width"           , "46"  ), _
  Array("Wio"                  , "46"  ), _
  Array("Lio_one"              , "250" ), _
  Array("Lio_two"              , "100" ), _
  Array("Lio_three"            , "150" ), _
  Array("Lio_four"             , "80"  ), _
  Array("Lio_five"             , "60"  ), _
  Array("Lio_six"              , "50"  ), _
  Array("Lio_seven"            , "100" ), _
  Array("Lio_eight"            , "120" ), _
  Array("Lio_nine"             , "60"  ), _
  Array("Lio_ten"              , "180" ), _
  Array("Lio_eleven"           , "130" ), _

  ' — keep‑out  —
  Array("Edge_wall_pushback"   , "40"  ), _

  ' — sweeping convenience —
  Array("ResonatorDeltaL"      , "0"   ) _                      ' global trim
)

Dim v
For Each v In vars
  oDesign.ChangeProperty Array("NAME:AllTabs",Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers","LocalVariables"), _
    Array("NAME:NewProps",Array("NAME:"&v(0),"PropType:=","VariableProp","UserDef:=",true,"Value:=",v(1)))))
Next

'------------------------------------------------------------------------------
' 4.  HELPER SUB FOR RESONATORS
'------------------------------------------------------------------------------
Sub CreateResonator(resNum, resType, x0, y0, z0, orientDeg, onTopLayer)
  Dim name, pts, slotW, slotL
  name = "RES"&CStr(resNum)
  If resType="A" Then
     Dim side : side = "ResA_side"
     slotW = "ResA_slot_W"
     slotL = "ResA_slot_L"
     ' square outline (centre justified)
     pts = Array( _                    '12‑point polyline (clockwise)
       P(x0 & "-"&side&"/2", y0 & "-"&side&"/2"), _
       P(x0 & "+" &side&"/2", y0 & "-"&side&"/2"), _
       P(x0 & "+" &side&"/2", y0 & "+" &side&"/2"), _
       P(x0 & "-" &side&"/2", y0 & "+" &side&"/2"), _
       P(x0 & "-" &side&"/2", y0 & "-" &side&"/2") )
     ' slot = simple rectangle  
     oEditor.CreateBox Array("NAME:BoxParameters","XPosition:=", _
         x0 & "-"&slotW&"/2","YPosition:=",y0&"-"&slotL&"/2","ZPosition:=",z0, _
         "XSize:=",slotW,"YSize:=",slotL,"ZSize:=","Cond_Trace"), _
         Array("NAME:Attributes","Name:=",name&"_slot","Flags:=","","Color:=","(0 0 255)", _
         "MaterialValue:=","\"air\"","SolveInside:=",true)
  Else
     ' type B (rectangle with twin slots)
     Dim w : w = "ResB_W" : Dim h : h = "ResB_H"
     pts = Array( _
       P(x0 & "-"&w&"/2", y0 & "-"&h&"/2"), _
       P(x0 & "+" &w&"/2", y0 & "-"&h&"/2"), _
       P(x0 & "+" &w&"/2", y0 & "+" &h&"/2"), _
       P(x0 & "-" &w&"/2", y0 & "+" &h&"/2"), _
       P(x0 & "-" &w&"/2", y0 & "-"&h&"/2") )
     ' two slots
     Dim sep : sep = "ResB_slot_sep"
     oEditor.CreateBox Array("NAME:BoxParameters","XPosition:=", _
         x0 & "-"&"ResB_slot_W"&"/2","YPosition:=",y0&"-"&sep&"/2 - ResA_slot_L/2","ZPosition:=",z0, _
         "XSize:=","ResB_slot_W","YSize:=","ResA_slot_L","ZSize:=","Cond_Trace"), _
         Array("NAME:Attributes","Name:=",name&"_slot1","Flags:=","","Color:=","(0 0 255)", _
         "MaterialValue:=","\"air\"","SolveInside:=",true)
     oEditor.CreateBox Array("NAME:BoxParameters","XPosition:=", _
         x0 & "-"&"ResB_slot_W"&"/2","YPosition:=",y0&"+"&sep&"/2","ZPosition:=",z0, _
         "XSize:=","ResB_slot_W","YSize:=","ResA_slot_L","ZSize:=","Cond_Trace"), _
         Array("NAME:Attributes","Name:=",name&"_slot2","Flags:=","","Color:=","(0 0 255)", _
         "MaterialValue:=","\"air\"","SolveInside:=",true)
  End If

  ' draw outside trace as polyline & sweep
  Dim polyPts, i: ReDim polyPts(UBound(pts))
  For i = 0 To UBound(pts): polyPts(i)=pts(i): Next
  oEditor.CreatePolyline Array( _
     "NAME:PolylineParameters","IsPolylineCovered:=",true,"IsPolylineClosed:=",true, _
     Array("NAME:PolylinePoints"), polyPts, _
     Array("NAME:PolylineSegments", Array("NAME:PLSegment","SegmentType:=","Line","StartIndex:=",0,"NoOfPoints:=",UBound(polyPts)+1) ), _
     Array("NAME:PolylineXSection","XSectionType:=","None","XSectionOrient:=","Auto"), _
     Array("NAME:Attributes","Name:=",name,"Color:=","(255 128 0)","MaterialValue:=","\"Copper_HP\"") )

  ' sweep to thickness
  Dim sel : sel = name
  oEditor.SweepAlongVector Array("NAME:Selections","Selections:=",sel,"NewPartsModelFlag:=","Model"), _
       Array("NAME:VectorSweepParameters","DraftAngle:=","0deg","DraftType:=","Round", _
             "DraftAutoSymm:=",false,"CheckFaceFaceIntersection:=",false, _
             "SweepVectorX:=","0mil","SweepVectorY:=","0mil","SweepVectorZ:=","Cond_Trace")
  ' subtract slots
  If resType="A" Then
     oEditor.Subtract Array("NAME:Selections","Blank Parts:=",sel,"Tool Parts:=",name&"_slot"), _
                      Array("NAME:SubtractParameters","KeepOriginals:=",false)
  Else
     oEditor.Subtract Array("NAME:Selections","Blank Parts:=",sel,"Tool Parts:=", _
                      name&"_slot1&,"&name&"_slot2"),Array("NAME:SubtractParameters","KeepOriginals:=",false)
  End If

  ' orientation (rotate) & layer placement
  If orientDeg<>0 Then _
     oEditor.Rotate Array("NAME:Selections","Selections:=",sel),Array("NAME:RotateParameters","RotateAxis:=","Z","RotateAngle:=",CStr(orientDeg)&"deg","RotatePoint:=",Array(x0,y0,z0))
  If onTopLayer=False Then _
     oEditor.Move Array("NAME:Selections","Selections:=",sel),Array("NAME:TranslateParameters","TranslateVector:=","0mil,0mil,-(" & "Prepreg_H/2" & ")")
End Sub

Function P(x,y): P = Array("NAME:PLPoint","X:=",x,"Y:=",y,"Z:=","0mil"): End Function

'------------------------------------------------------------------------------
' 5.  STACK‑UP  (all depend on variables)
'------------------------------------------------------------------------------
Sub MakeStack
  Dim xSize : xSize="4000"        ' enough margin; parametric bounding box can be added
  Dim ySize : ySize="3000"
  oEditor.CreateBox Array("NAME:BoxParameters","XPosition:=","-Edge_wall_pushback","YPosition:=","-Edge_wall_pushback", _
        "ZPosition:=","-Core_H2","XSize:=",xSize,"YSize:=",ySize,"ZSize:=","Core_H2"), _
        Array("NAME:Attributes","Name:=","CORE_BOT","MaterialValue:=","\"Rogers3010\"","Color:=","(200 200 200)")
  oEditor.CreateBox Array("NAME:BoxParameters","XPosition:=","-Edge_wall_pushback","YPosition:=","-Edge_wall_pushback", _
        "ZPosition:=","0","XSize:=",xSize,"YSize:=",ySize,"ZSize:=","Prepreg_H"), _
        Array("NAME:Attributes","Name:=","PREPREG","MaterialValue:=","\"Rogers2929\"","Color:=","(180 220 180)")
  oEditor.CreateBox Array("NAME:BoxParameters","XPosition:=","-Edge_wall_pushback","YPosition:=","-Edge_wall_pushback", _
        "ZPosition:=","Prepreg_H","XSize:=",xSize,"YSize:=",ySize,"ZSize:=","Core_H1"), _
        Array("NAME:Attributes","Name:=","CORE_TOP","MaterialValue:=","\"Rogers3010\"","Color:=","(200 200 200)")

  ' ground planes
  oEditor.CreateBox Array("NAME:BoxParameters","XPosition:=","-Edge_wall_pushback","YPosition:=","-Edge_wall_pushback", _
        "ZPosition:=","-Core_H2-Cond_Bot","XSize:=",xSize,"YSize:=",ySize,"ZSize:=","Cond_Bot"), _
        Array("NAME:Attributes","Name:=","GND_BOT","MaterialValue:=","\"Copper_HP\"","Color:=","(128 128 0)")
  oEditor.CreateBox Array("NAME:BoxParameters","XPosition:=","-Edge_wall_pushback","YPosition:=","-Edge_wall_pushback", _
        "ZPosition:=","Prepreg_H+Core_H1","XSize:=",xSize,"YSize:=",ySize,"ZSize:=","Cond_Top"), _
        Array("NAME:Attributes","Name:=","GND_TOP","MaterialValue:=","\"Copper_HP\"","Color:=","(128 128 0)")
End Sub
MakeStack

'------------------------------------------------------------------------------
' 6.  CREATE RESONATORS  (mirrored pairs)
'------------------------------------------------------------------------------
Dim cx: cx="0"     ' centre line X
Dim y0: y0="0"

' left side  (negative X)
CreateResonator 1,"A","-3*Res_x_spacing*XYSpacing_scale",     y0,                  "0mil",0, True
CreateResonator 2,"B","-2*Res_x_spacing*XYSpacing_scale",     "y0-Res_y_spacing",  "0mil",0, False
CreateResonator 3,"B","-1*Res_x_spacing*XYSpacing_scale",     "y0+Res_y_spacing",  "0mil",0, True
CreateResonator 4,"A","0",                                    "y0",                "0mil",0, False

' right side is mirror image about X=0
CreateResonator 5,"A","(3*Res_x_spacing*XYSpacing_scale)",    y0,                  "0mil",180,False
CreateResonator 6,"B","(2*Res_x_spacing*XYSpacing_scale)",    "y0+Res_y_spacing",  "0mil",180,True
CreateResonator 7,"B","(1*Res_x_spacing*XYSpacing_scale)",    "y0-Res_y_spacing",  "0mil",180,False
CreateResonator 8,"A","0",                                    "y0",                "0mil",180,True

'------------------------------------------------------------------------------
' 7.  LAUNCH / FEED NETWORK (eleven segments, left & right)
'------------------------------------------------------------------------------
Sub Seg(xlen, ylen, name)
  oEditor.CreatePolyline Array("NAME:PolylineParameters","IsPolylineCovered:=",true,"IsPolylineClosed:=",false, _
      Array("NAME:PolylinePoints",P("0","0"),P(xlen,ylen)), _
      Array("NAME:PolylineSegments",Array("NAME:PLSegment","SegmentType:=","Line","StartIndex:=",0,"NoOfPoints:=",2)), _
      Array("NAME:PolylineXSection","XSectionType:=","None"), _
      Array("NAME:Attributes","Name:=",name,"MaterialValue:=","\"Copper_HP\""))
  oEditor.SweepAlongVector Array("NAME:Selections","Selections:=",name,"NewPartsModelFlag:=","Model"), _
        Array("NAME:VectorSweepParameters","SweepVectorX:=","0","SweepVectorY:=","0","SweepVectorZ:=","Cond_Trace")
End Sub

' Example feed on left – user can alter path easily via variables
Seg "-Lio_one","0","Feed1_seg1"
Seg "0","-Lio_two","Feed1_seg2"
Seg "-Lio_three","-Lio_four","Feed1_seg3"
' … continue until Feed1_seg11
' Mirror to right side with Move/Scale later or repeat manually.

'------------------------------------------------------------------------------
' 8.  PORTS
'------------------------------------------------------------------------------
Dim oBnd : Set oBnd = oDesign.GetModule("BoundarySetup")
oEditor.CreateRectangle Array("NAME:RectangleParameters","IsCovered:=",true,"XStart:=","-Lio_one","YStart:=","Wio/2-Port_Width/2","ZStart:=","0mil","Width:=","Port_Width","Height:=","Cond_Trace","WhichAxis:=","X"), _
      Array("NAME:Attributes","Name:=","Port1_Obj","MaterialValue:=","\"Copper_HP\"")
oBnd.AssignWavePort Array("NAME:Port1","Objects:=",Array("Port1_Obj"),"NumModes:=",1,"RenormalizeAllTerminals:=",true, _
      "UseLineAlignment:=",false,"DoDeembed:=",true,"DeembedDist:=","0mil", _
      Array("NAME:Modes",Array("NAME:Mode1","ModeNum:=",1,"UseIntLine:=",true, _
      Array("NAME:IntLine","Start:=",Array("-Lio_one","0","0"),"End:=",Array("-Lio_one","0","Cond_Trace") ),"CharImp:=","Zpi")))

' Second port on right
oEditor.CreateRectangle Array("NAME:RectangleParameters","IsCovered:=",true,"XStart:=","Lio_one","YStart:=","Wio/2-Port_Width/2","ZStart:=","0mil","Width:=","Port_Width","Height:=","Cond_Trace","WhichAxis:=","X"), _
      Array("NAME:Attributes","Name:=","Port2_Obj","MaterialValue:=","\"Copper_HP\"")
oBnd.AssignWavePort Array("NAME:Port2","Objects:=",Array("Port2_Obj"),"NumModes:=",1,"RenormalizeAllTerminals:=",true, _
      "UseLineAlignment:=",false,"DoDeembed:=",true,"DeembedDist:=","0mil", _
      Array("NAME:Modes",Array("NAME:Mode1","ModeNum:=",1,"UseIntLine:=",true, _
      Array("NAME:IntLine","Start:=",Array("Lio_one","0","0"),"End:=",Array("Lio_one","0","Cond_Trace") ),"CharImp:=","Zpi")))

'------------------------------------------------------------------------------
' 9.  ANALYSIS SET‑UPS
'------------------------------------------------------------------------------
Set oSetup = oDesign.GetModule("AnalysisSetup")
Dim fc : fc="fc_GHz & \"GHz\""
oSetup.InsertSetup "HfssDriven", Array("NAME:Setup1","Frequency:=",fc,"MaxDeltaS:=",0.01,"MaximumPasses:=",20,"PercentRefinement:=",15)
oSetup.InsertFrequencySweep "Setup1", Array("NAME:SweepFast","IsEnabled:=",true,"SetupType:=","LinearStep", _
      "StartValue:=","StartFreq_GHz&\"GHz\"","StopValue:=","StopFreq_GHz&\"GHz\"","StepSize:=","0.02GHz","Type:=","Fast")

' two mesh‑linked single‑point solves
oSetup.InsertSetup "HfssDriven", Array("NAME:Setup_fc+","Frequency:=","fc_GHz+0.3&\"GHz\"","MaximumPasses:=",4, _
      "MeshLink:=",Array("NAME:MeshLink","Project:=","This Project","Design:=","This Design","Soln:=","Setup1 : LastAdaptive") )
oSetup.InsertSetup "HfssDriven", Array("NAME:Setup_fc-","Frequency:=","fc_GHz-0.3&\"GHz\"","MaximumPasses:=",4, _
      "MeshLink:=",Array("NAME:MeshLink","Project:=","This Project","Design:=","This Design","Soln:=","Setup_fc+ : LastAdaptive") )

'------------------------------------------------------------------------------
' 10.  REPORT TEMPLATE
'------------------------------------------------------------------------------
Dim oRpt : Set oRpt = oDesign.GetModule("ReportSetup")
oRpt.CreateReport "S‑parameters","Modal Solution Data","Rectangular Plot","Setup1 : SweepFast", _
      Array("Freq"),Array("X Component:=","Freq","Y Component:=","dB(S(Port1,Port2)) dB"), _
      Array("Context:=","", "plot_singularities:=",false)

MsgBox "8‑pole stripline filter generated successfully!"
