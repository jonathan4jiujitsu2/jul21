'***************************************************************************
' HFSS VBScript for 8-Pole Multi-Layer Stripline Filter
' Based on mirrored resonator layout with complex launch structures
' NO VIAS - Pure stripline design with multiple resonator types
'***************************************************************************

'Initialize variables
Dim oAnsoftApp, oDesktop, oProject, oDesign, oEditor, oModule, oDefinitionManager

'Set project objects
Set oAnsoftApp = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow
Set oProject = oDesktop.NewProject
Set oProject = oDesktop.GetActiveProject
oProject.InsertDesign "HFSS", "MultiLayerStriplineFilter_8Pole", "DrivenModal", ""
Set oDesign = oProject.GetActiveDesign
Set oEditor = oDesign.SetActiveEditor("3D Modeler")

'Set units to mils
oEditor.SetModelUnits Array("NAME:Units Parameter", "Units:=", "mil", "Rescale:=", false)

'***************************************************************************
' Material Definitions
'***************************************************************************

Set oDefinitionManager = oProject.GetDefinitionManager()

'Create Rogers 3010 material (εr = 10.2)
oDefinitionManager.AddMaterial Array("NAME:Rogers3010", _
    "CoordinateSystemType:=", "Cartesian", _
    Array("NAME:AttachedData"), _
    "permittivity:=", "10.2", _
    "dielectric_loss_tangent:=", "0.0023")

'Create 2929 prepreg material
oDefinitionManager.AddMaterial Array("NAME:Prepreg2929", _
    "CoordinateSystemType:=", "Cartesian", _
    Array("NAME:AttachedData"), _
    "permittivity:=", "2.94", _
    "dielectric_loss_tangent:=", "0.003")

'Create high-performance copper
oDefinitionManager.AddMaterial Array("NAME:HighPerfCopper", _
    "CoordinateSystemType:=", "Cartesian", _
    Array("NAME:AttachedData"), _
    Array("NAME:ModifierData"), _
    "conductivity:=", "5.8e7")

'***************************************************************************
' Design Parameters
'***************************************************************************

'Solution Setup Parameters
Dim solution_freq_value, start_freq, end_freq, sweep_type
solution_freq_value = 15  'GHz - Center frequency
start_freq = "10GHz"
end_freq = "20GHz"
sweep_type = "Interpolating"
Dim max_delta_s, num_passes, percent_refine
max_delta_s = 0.01
num_passes = 40
percent_refine = 20

'Layer Stack Parameters
Dim layer_h_core1, layer_h_prepreg, layer_h_core2, cond_h
layer_h_core1 = 10     'mil - Top core layer
layer_h_prepreg = 5    'mil - Prepreg layer
layer_h_core2 = 10     'mil - Bottom core layer
cond_h = 0.5          'mil - Conductor thickness

'Resonator Type A (Square) Parameters
Dim resA_outer_size, resA_slot_size, resA_slot_offset
resA_outer_size = 80   'mil - Outer dimension of square resonator
resA_slot_size = 30    'mil - Central slot size
resA_slot_offset = 25  'mil - Slot offset from edge

'Resonator Type B (Rectangular) Parameters
Dim resB_width, resB_height, resB_slot_width, resB_slot_height, resB_gap
resB_width = 70        'mil - Width of rectangular resonator
resB_height = 90       'mil - Height of rectangular resonator
resB_slot_width = 20   'mil - Slot width
resB_slot_height = 35  'mil - Slot height
resB_gap = 5           'mil - Gap between dual slots

'Spacing Parameters
Dim res_x_spacing, res_y_spacing, res_z_spacing
res_x_spacing = 25     'mil - X spacing between resonators
res_y_spacing = 35     'mil - Y spacing between resonators
res_z_spacing = layer_h_prepreg  'Z spacing between layers

'Edge Clearance Parameters
Dim edge_wall_pullback, edge_wall_pullback2, edge_wall_pullback3, edge_wall_pullback4, edge_wall_pullback5
edge_wall_pullback = 20    'mil
edge_wall_pullback2 = 15   'mil
edge_wall_pullback3 = 25   'mil
edge_wall_pullback4 = 18   'mil
edge_wall_pullback5 = 22   'mil

'Launch Line Parameters
Dim Lio, Wio, Port_Width, Launch_width, Launch_spacing
Lio = 100              'mil - Port extension length
Wio = 8                'mil - Launch line width
Port_Width = 50        'mil - Port width
Launch_width = 4       'mil - Launch trace width
Launch_spacing = 10    'mil - Launch spacing

'Complex Launch Segments (from images)
Dim Lio_one, Lio_two, Lio_three, Lio_four, Lio_five, Lio_six
Dim Lio_seven, Lio_eight, Lio_nine, Lio_ten, Lio_eleven
Dim Lio_two_two, Wio_two
Lio_one = 80
Lio_two = 60
Lio_three = 40
Lio_four = 30
Lio_five = 35
Lio_six = 45
Lio_seven = 38
Lio_eight = 42
Lio_nine = 55
Lio_ten = 65
Lio_eleven = 75
Lio_two_two = 50
Wio_two = 6

'Board dimensions
Dim board_width, board_length
board_width = 600      'mil
board_length = 800     'mil

'Calculate layer Z positions
Dim z_gnd_bottom, z_core2_bottom, z_core2_top, z_prepreg_bottom, z_prepreg_top
Dim z_core1_bottom, z_core1_top, z_gnd_top
z_gnd_bottom = 0
z_core2_bottom = cond_h
z_core2_top = z_core2_bottom + layer_h_core2
z_prepreg_bottom = z_core2_top
z_prepreg_top = z_prepreg_bottom + layer_h_prepreg
z_core1_bottom = z_prepreg_top
z_core1_top = z_core1_bottom + layer_h_core1
z_gnd_top = z_core1_top

'***************************************************************************
' Create Design Variables
'***************************************************************************

'Create frequency variables
oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:CenterFreq", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(solution_freq_value) & "GHz"))))

'Create layer height variables
oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:layer_h_core1", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(layer_h_core1) & "mil"))))

oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:layer_h_prepreg", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(layer_h_prepreg) & "mil"))))

oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:layer_h_core2", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(layer_h_core2) & "mil"))))

oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:cond_h", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(cond_h) & "mil"))))

'Create resonator dimension variables
oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:resA_outer_size", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(resA_outer_size) & "mil"))))

oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:resA_slot_size", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(resA_slot_size) & "mil"))))

oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:resB_width", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(resB_width) & "mil"))))

oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:resB_height", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(resB_height) & "mil"))))

'Create spacing variables
oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:res_x_spacing", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(res_x_spacing) & "mil"))))

oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:res_y_spacing", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(res_y_spacing) & "mil"))))

'Create launch line variables
oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:Lio", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(Lio) & "mil"))))

oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:Wio", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(Wio) & "mil"))))

oDesign.ChangeProperty Array("NAME:AllTabs", _
    Array("NAME:LocalVariableTab", _
    Array("NAME:PropServers", "LocalVariables"), _
    Array("NAME:NewProps", _
    Array("NAME:Port_Width", "PropType:=", "VariableProp", _
    "UserDef:=", true, "Value:=", CStr(Port_Width) & "mil"))))

'***************************************************************************
' Create Layer Stack
'***************************************************************************

'Bottom ground plane
oEditor.CreateBox Array("NAME:BoxParameters", _
    "XPosition:=", "-" & CStr(board_length/2) & "mil", _
    "YPosition:=", "-" & CStr(board_width/2) & "mil", _
    "ZPosition:=", CStr(z_gnd_bottom) & "mil", _
    "XSize:=", CStr(board_length) & "mil", _
    "YSize:=", CStr(board_width) & "mil", _
    "ZSize:=", "cond_h"), _
    Array("NAME:Attributes", "Name:=", "GND_Bottom", "Flags:=", "", _
    "Color:=", "(132 132 193)", "Transparency:=", 0, _
    "PartCoordinateSystem:=", "Global", _
    "MaterialValue:=", Chr(34) & "HighPerfCopper" & Chr(34), _
    "SolveInside:=", false)

'Bottom core layer (Rogers 3010)
oEditor.CreateBox Array("NAME:BoxParameters", _
    "XPosition:=", "-" & CStr(board_length/2) & "mil", _
    "YPosition:=", "-" & CStr(board_width/2) & "mil", _
    "ZPosition:=", CStr(z_core2_bottom) & "mil", _
    "XSize:=", CStr(board_length) & "mil", _
    "YSize:=", CStr(board_width) & "mil", _
    "ZSize:=", "layer_h_core2"), _
    Array("NAME:Attributes", "Name:=", "Core2_Bottom", "Flags:=", "", _
    "Color:=", "(0 128 0)", "Transparency:=", 0.7, _
    "PartCoordinateSystem:=", "Global", _
    "MaterialValue:=", Chr(34) & "Rogers3010" & Chr(34), _
    "SolveInside:=", true)

'Prepreg layer
oEditor.CreateBox Array("NAME:BoxParameters", _
    "XPosition:=", "-" & CStr(board_length/2) & "mil", _
    "YPosition:=", "-" & CStr(board_width/2) & "mil", _
    "ZPosition:=", CStr(z_prepreg_bottom) & "mil", _
    "XSize:=", CStr(board_length) & "mil", _
    "YSize:=", CStr(board_width) & "mil", _
    "ZSize:=", "layer_h_prepreg"), _
    Array("NAME:Attributes", "Name:=", "Prepreg_Middle", "Flags:=", "", _
    "Color:=", "(255 255 0)", "Transparency:=", 0.7, _
    "PartCoordinateSystem:=", "Global", _
    "MaterialValue:=", Chr(34) & "Prepreg2929" & Chr(34), _
    "SolveInside:=", true)

'Top core layer (Rogers 3010)
oEditor.CreateBox Array("NAME:BoxParameters", _
    "XPosition:=", "-" & CStr(board_length/2) & "mil", _
    "YPosition:=", "-" & CStr(board_width/2) & "mil", _
    "ZPosition:=", CStr(z_core1_bottom) & "mil", _
    "XSize:=", CStr(board_length) & "mil", _
    "YSize:=", CStr(board_width) & "mil", _
    "ZSize:=", "layer_h_core1"), _
    Array("NAME:Attributes", "Name:=", "Core1_Top", "Flags:=", "", _
    "Color:=", "(0 128 0)", "Transparency:=", 0.7, _
    "PartCoordinateSystem:=", "Global", _
    "MaterialValue:=", Chr(34) & "Rogers3010" & Chr(34), _
    "SolveInside:=", true)

'Top ground plane
oEditor.CreateBox Array("NAME:BoxParameters", _
    "XPosition:=", "-" & CStr(board_length/2) & "mil", _
    "YPosition:=", "-" & CStr(board_width/2) & "mil", _
    "ZPosition:=", CStr(z_gnd_top) & "mil", _
    "XSize:=", CStr(board_length) & "mil", _
    "YSize:=", CStr(board_width) & "mil", _
    "ZSize:=", "cond_h"), _
    Array("NAME:Attributes", "Name:=", "GND_Top", "Flags:=", "", _
    "Color:=", "(132 132 193)", "Transparency:=", 0, _
    "PartCoordinateSystem:=", "Global", _
    "MaterialValue:=", Chr(34) & "HighPerfCopper" & Chr(34), _
    "SolveInside:=", false)

'***************************************************************************
' Subroutine to Create Resonators
'***************************************************************************

Sub CreateResonator(res_num, res_type, x_pos, y_pos, z_pos, orientation)
    Dim res_name
    res_name = "Resonator" & res_num
    
    If res_type = "A" Then
        'Create Type A resonator (Square with single slot)
        
        'Main resonator body
        oEditor.CreateBox Array("NAME:BoxParameters", _
            "XPosition:=", CStr(x_pos) & "mil", _
            "YPosition:=", CStr(y_pos) & "mil", _
            "ZPosition:=", CStr(z_pos) & "mil", _
            "XSize:=", "resA_outer_size", _
            "YSize:=", "resA_outer_size", _
            "ZSize:=", "cond_h"), _
            Array("NAME:Attributes", "Name:=", res_name & "_Body", _
            "Flags:=", "", "Color:=", "(255 128 0)", _
            "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
            "MaterialValue:=", Chr(34) & "HighPerfCopper" & Chr(34), _
            "SolveInside:=", false)
        
        'Create slot
        oEditor.CreateBox Array("NAME:BoxParameters", _
            "XPosition:=", CStr(x_pos + resA_slot_offset) & "mil", _
            "YPosition:=", CStr(y_pos + resA_slot_offset) & "mil", _
            "ZPosition:=", CStr(z_pos) & "mil", _
            "XSize:=", "resA_slot_size", _
            "YSize:=", "resA_slot_size", _
            "ZSize:=", "cond_h"), _
            Array("NAME:Attributes", "Name:=", res_name & "_Slot", _
            "Flags:=", "", "Color:=", "(255 0 0)", _
            "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
            "MaterialValue:=", Chr(34) & "vacuum" & Chr(34), _
            "SolveInside:=", true)
        
        'Subtract slot from body
        oEditor.Subtract Array("NAME:Selections", _
            "Blank Parts:=", res_name & "_Body", _
            "Tool Parts:=", res_name & "_Slot"), _
            Array("NAME:SubtractParameters", "KeepOriginals:=", false)
            
    ElseIf res_type = "B" Then
        'Create Type B resonator (Rectangular with dual slots)
        
        'Main resonator body
        oEditor.CreateBox Array("NAME:BoxParameters", _
            "XPosition:=", CStr(x_pos) & "mil", _
            "YPosition:=", CStr(y_pos) & "mil", _
            "ZPosition:=", CStr(z_pos) & "mil", _
            "XSize:=", "resB_width", _
            "YSize:=", "resB_height", _
            "ZSize:=", "cond_h"), _
            Array("NAME:Attributes", "Name:=", res_name & "_Body", _
            "Flags:=", "", "Color:=", "(0 255 0)", _
            "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
            "MaterialValue:=", Chr(34) & "HighPerfCopper" & Chr(34), _
            "SolveInside:=", false)
        
        'Create first slot
        oEditor.CreateBox Array("NAME:BoxParameters", _
            "XPosition:=", CStr(x_pos + (resB_width - resB_slot_width)/2) & "mil", _
            "YPosition:=", CStr(y_pos + 10) & "mil", _
            "ZPosition:=", CStr(z_pos) & "mil", _
            "XSize:=", "resB_slot_width", _
            "YSize:=", "resB_slot_height", _
            "ZSize:=", "cond_h"), _
            Array("NAME:Attributes", "Name:=", res_name & "_Slot1", _
            "Flags:=", "", "Color:=", "(255 0 0)", _
            "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
            "MaterialValue:=", Chr(34) & "vacuum" & Chr(34), _
            "SolveInside:=", true)
        
        'Create second slot
        oEditor.CreateBox Array("NAME:BoxParameters", _
            "XPosition:=", CStr(x_pos + (resB_width - resB_slot_width)/2) & "mil", _
            "YPosition:=", CStr(y_pos + resB_height - resB_slot_height - 10) & "mil", _
            "ZPosition:=", CStr(z_pos) & "mil", _
            "XSize:=", "resB_slot_width", _
            "YSize:=", "resB_slot_height", _
            "ZSize:=", "cond_h"), _
            Array("NAME:Attributes", "Name:=", res_name & "_Slot2", _
            "Flags:=", "", "Color:=", "(255 0 0)", _
            "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
            "MaterialValue:=", Chr(34) & "vacuum" & Chr(34), _
            "SolveInside:=", true)
        
        'Subtract slots from body
        oEditor.Subtract Array("NAME:Selections", _
            "Blank Parts:=", res_name & "_Body", _
            "Tool Parts:=", res_name & "_Slot1," & res_name & "_Slot2"), _
            Array("NAME:SubtractParameters", "KeepOriginals:=", false)
    End If
    
    'Apply rotation if needed
    If orientation <> 0 Then
        oEditor.Rotate Array("NAME:Selections", _
            "Selections:=", res_name & "_Body", _
            "NewPartsModelFlag:=", "Model"), _
            Array("NAME:RotateParameters", _
            "RotateAxis:=", "Z", _
            "RotateAngle:=", CStr(orientation) & "deg", _
            "RotateCenter:=", Array(CStr(x_pos + resA_outer_size/2) & "mil", _
                CStr(y_pos + resA_outer_size/2) & "mil", _
                CStr(z_pos) & "mil"))
    End If
End Sub

'***************************************************************************
' Create Resonators
'***************************************************************************

'Define base positions
Dim base_x, base_y
base_x = -150  'Starting X position
base_y = -100  'Starting Y position

'Left side resonators
'Resonator 1 - Type A on lower layer
CreateResonator 1, "A", base_x, base_y, z_core2_bottom + cond_h, 0

'Resonator 2 - Type B on upper layer
CreateResonator 2, "B", base_x, base_y + resA_outer_size + res_y_spacing, _
    z_prepreg_top + cond_h, 0

'Resonator 3 - Type B on lower layer
CreateResonator 3, "B", base_x + resA_outer_size + res_x_spacing, base_y, _
    z_core2_bottom + cond_h, 0

'Resonator 4 - Type A on upper layer
CreateResonator 4, "A", base_x + resA_outer_size + res_x_spacing, _
    base_y + resB_height + res_y_spacing, z_prepreg_top + cond_h, 0

'Right side resonators (mirrored)
Dim mirror_x_offset
mirror_x_offset = 300  'Offset for mirroring

'Resonator 5 - Type A on upper layer (mirror of Res4)
CreateResonator 5, "A", base_x + mirror_x_offset, _
    base_y + resB_height + res_y_spacing, z_prepreg_top + cond_h, 0

'Resonator 6 - Type B on lower layer (mirror of Res3)
CreateResonator 6, "B", base_x + mirror_x_offset + resA_outer_size - resB_width, _
    base_y, z_core2_bottom + cond_h, 0

'Resonator 7 - Type B on upper layer (mirror of Res2)
CreateResonator 7, "B", base_x + mirror_x_offset + resA_outer_size - resB_width, _
    base_y + resA_outer_size + res_y_spacing, z_prepreg_top + cond_h, 0

'Resonator 8 - Type A on lower layer (mirror of Res1)
CreateResonator 8, "A", base_x + mirror_x_offset, base_y, z_core2_bottom + cond_h, 0

'***************************************************************************
' Create Complex Launch Structure
'***************************************************************************

'Left side launch
Dim launch_points_left(11, 2)  'Array for launch line coordinates
launch_points_left(0, 0) = -Lio - 50
launch_points_left(0, 1) = base_y + resA_outer_size/2
launch_points_left(1, 0) = base_x - 30
launch_points_left(1, 1) = base_y + resA_outer_size/2
launch_points_left(2, 0) = base_x - 20
launch_points_left(2, 1) = base_y + resA_outer_size/2 - 10
launch_points_left(3, 0) = base_x - 10
launch_points_left(3, 1) = base_y + resA_outer_size/2 - 20
launch_points_left(4, 0) = base_x
launch_points_left(4, 1) = base_y + resA_outer_size/2

'Create left launch line using polyline
oEditor.CreatePolyline Array("NAME:PolylineParameters", _
    "IsPolylineCovered:=", true, "IsPolylineClosed:=", false, _
    Array("NAME:PolylinePoints", _
        Array("NAME:PLPoint", "X:=", CStr(launch_points_left(0, 0)) & "mil", _
            "Y:=", CStr(launch_points_left(0, 1)) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(launch_points_left(1, 0)) & "mil", _
            "Y:=", CStr(launch_points_left(1, 1)) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(launch_points_left(2, 0)) & "mil", _
            "Y:=", CStr(launch_points_left(2, 1)) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(launch_points_left(3, 0)) & "mil", _
            "Y:=", CStr(launch_points_left(3, 1)) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(launch_points_left(4, 0)) & "mil", _
            "Y:=", CStr(launch_points_left(4, 1)) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil")), _
    Array("NAME:PolylineSegments", _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 0, "NoOfPoints:=", 2), _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 1, "NoOfPoints:=", 2), _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 2, "NoOfPoints:=", 2), _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 3, "NoOfPoints:=", 2)), _
    Array("NAME:PolylineXSection", "XSectionType:=", "Rectangle", _
        "XSectionOrient:=", "Auto", _
        "XSectionWidth:=", "Wio", _
        "XSectionTopWidth:=", "0mil", _
        "XSectionHeight:=", "cond_h", _
        "XSectionNumSegments:=", "0", _
        "XSectionBendType:=", "Corner")), _
    Array("NAME:Attributes", "Name:=", "LaunchLine_Left", _
        "Flags:=", "", "Color:=", "(255 128 0)", _
        "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
        "MaterialValue:=", Chr(34) & "HighPerfCopper" & Chr(34), _
        "SolveInside:=", false)

'Right side launch (mirrored)
Dim launch_points_right(11, 2)
launch_points_right(0, 0) = base_x + mirror_x_offset + resA_outer_size
launch_points_right(0, 1) = base_y + resA_outer_size/2
launch_points_right(1, 0) = base_x + mirror_x_offset + resA_outer_size + 10
launch_points_right(1, 1) = base_y + resA_outer_size/2 - 20
launch_points_right(2, 0) = base_x + mirror_x_offset + resA_outer_size + 20
launch_points_right(2, 1) = base_y + resA_outer_size/2 - 10
launch_points_right(3, 0) = base_x + mirror_x_offset + resA_outer_size + 30
launch_points_right(3, 1) = base_y + resA_outer_size/2
launch_points_right(4, 0) = base_x + mirror_x_offset + resA_outer_size + Lio + 50
launch_points_right(4, 1) = base_y + resA_outer_size/2

'Create right launch line
oEditor.CreatePolyline Array("NAME:PolylineParameters", _
    "IsPolylineCovered:=", true, "IsPolylineClosed:=", false, _
    Array("NAME:PolylinePoints", _
        Array("NAME:PLPoint", "X:=", CStr(launch_points_right(0, 0)) & "mil", _
            "Y:=", CStr(launch_points_right(0, 1)) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(launch_points_right(1, 0)) & "mil", _
            "Y:=", CStr(launch_points_right(1, 1)) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(launch_points_right(2, 0)) & "mil", _
            "Y:=", CStr(launch_points_right(2, 1)) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(launch_points_right(3, 0)) & "mil", _
            "Y:=", CStr(launch_points_right(3, 1)) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(launch_points_right(4, 0)) & "mil", _
            "Y:=", CStr(launch_points_right(4, 1)) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil")), _
    Array("NAME:PolylineSegments", _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 0, "NoOfPoints:=", 2), _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 1, "NoOfPoints:=", 2), _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 2, "NoOfPoints:=", 2), _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 3, "NoOfPoints:=", 2)), _
    Array("NAME:PolylineXSection", "XSectionType:=", "Rectangle", _
        "XSectionOrient:=", "Auto", _
        "XSectionWidth:=", "Wio", _
        "XSectionTopWidth:=", "0mil", _
        "XSectionHeight:=", "cond_h", _
        "XSectionNumSegments:=", "0", _
        "XSectionBendType:=", "Corner")), _
    Array("NAME:Attributes", "Name:=", "LaunchLine_Right", _
        "Flags:=", "", "Color:=", "(255 128 0)", _
        "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
        "MaterialValue:=", Chr(34) & "HighPerfCopper" & Chr(34), _
        "SolveInside:=", false)

'***************************************************************************
' Create Inter-resonator Coupling Lines
'***************************************************************************

'Coupling between Res1 and Res3
oEditor.CreatePolyline Array("NAME:PolylineParameters", _
    "IsPolylineCovered:=", true, "IsPolylineClosed:=", false, _
    Array("NAME:PolylinePoints", _
        Array("NAME:PLPoint", "X:=", CStr(base_x + resA_outer_size) & "mil", _
            "Y:=", CStr(base_y + resA_outer_size/2) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(base_x + resA_outer_size + res_x_spacing/2) & "mil", _
            "Y:=", CStr(base_y + resA_outer_size/2) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(base_x + resA_outer_size + res_x_spacing) & "mil", _
            "Y:=", CStr(base_y + resB_height/2) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil")), _
    Array("NAME:PolylineSegments", _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 0, "NoOfPoints:=", 2), _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 1, "NoOfPoints:=", 2)), _
    Array("NAME:PolylineXSection", "XSectionType:=", "Rectangle", _
        "XSectionOrient:=", "Auto", _
        "XSectionWidth:=", CStr(Launch_width) & "mil", _
        "XSectionTopWidth:=", "0mil", _
        "XSectionHeight:=", "cond_h", _
        "XSectionNumSegments:=", "0", _
        "XSectionBendType:=", "Corner")), _
    Array("NAME:Attributes", "Name:=", "Coupling_1_3", _
        "Flags:=", "", "Color:=", "(0 255 255)", _
        "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
        "MaterialValue:=", Chr(34) & "HighPerfCopper" & Chr(34), _
        "SolveInside:=", false)

'Coupling between Res2 and Res4 (vertical interlayer coupling)
oEditor.CreateBox Array("NAME:BoxParameters", _
    "XPosition:=", CStr(base_x + resA_outer_size + res_x_spacing + resB_width/2 - 2) & "mil", _
    "YPosition:=", CStr(base_y + resA_outer_size + res_y_spacing + resB_height) & "mil", _
    "ZPosition:=", CStr(z_prepreg_bottom) & "mil", _
    "XSize:=", "4mil", _
    "YSize:=", CStr(res_y_spacing) & "mil", _
    "ZSize:=", "layer_h_prepreg"), _
    Array("NAME:Attributes", "Name:=", "Coupling_2_4", _
    "Flags:=", "", "Color:=", "(255 0 255)", _
    "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
    "MaterialValue:=", Chr(34) & "HighPerfCopper" & Chr(34), _
    "SolveInside:=", false)

'Mirror coupling for right side
'Coupling between Res6 and Res8
oEditor.CreatePolyline Array("NAME:PolylineParameters", _
    "IsPolylineCovered:=", true, "IsPolylineClosed:=", false, _
    Array("NAME:PolylinePoints", _
        Array("NAME:PLPoint", "X:=", CStr(base_x + mirror_x_offset + resA_outer_size - resB_width) & "mil", _
            "Y:=", CStr(base_y + resB_height/2) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(base_x + mirror_x_offset - res_x_spacing/2) & "mil", _
            "Y:=", CStr(base_y + resA_outer_size/2) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil"), _
        Array("NAME:PLPoint", "X:=", CStr(base_x + mirror_x_offset) & "mil", _
            "Y:=", CStr(base_y + resA_outer_size/2) & "mil", _
            "Z:=", CStr(z_core2_bottom + cond_h) & "mil")), _
    Array("NAME:PolylineSegments", _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 0, "NoOfPoints:=", 2), _
        Array("NAME:PLSegment", "SegmentType:=", "Line", _
            "StartIndex:=", 1, "NoOfPoints:=", 2)), _
    Array("NAME:PolylineXSection", "XSectionType:=", "Rectangle", _
        "XSectionOrient:=", "Auto", _
        "XSectionWidth:=", CStr(Launch_width) & "mil", _
        "XSectionTopWidth:=", "0mil", _
        "XSectionHeight:=", "cond_h", _
        "XSectionNumSegments:=", "0", _
        "XSectionBendType:=", "Corner")), _
    Array("NAME:Attributes", "Name:=", "Coupling_6_8", _
        "Flags:=", "", "Color:=", "(0 255 255)", _
        "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
        "MaterialValue:=", Chr(34) & "HighPerfCopper" & Chr(34), _
        "SolveInside:=", false)

'***************************************************************************
' Create Ports
'***************************************************************************

'Port 1 - Left side
oEditor.CreateRectangle Array("NAME:RectangleParameters", _
    "IsCovered:=", true, _
    "XStart:=", CStr(launch_points_left(0, 0)) & "mil", _
    "YStart:=", CStr(launch_points_left(0, 1) - Port_Width/2) & "mil", _
    "ZStart:=", CStr(z_core2_bottom) & "mil", _
    "Width:=", "Port_Width", _
    "Height:=", CStr(z_core2_top - z_core2_bottom) & "mil", _
    "WhichAxis:=", "X"), _
    Array("NAME:Attributes", "Name:=", "Port1", _
    "Flags:=", "", "Color:=", "(0 0 255)", _
    "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", _
    "MaterialValue:=", Chr(34) & "vacuum" & Chr(34), _
    "SolveInside:=", true)

'Port 2 - Right side
oEditor.CreateRectangle Array("NAME:RectangleParameters", _
    "IsCovered:=", true, _
    "XStart:=", CStr(launch_points_right(4, 0)) & "mil", _
    "YStart:=", CStr(launch_points_right(4, 1) - Port_Width/2) & "mil", _
    "ZStart:=", CStr(z_core2_bottom) & "mil", _
    "Width:=", "Port_Width", _
    "Height:=", CStr(z_core2_top - z_core2_bottom) & "mil", _
    "WhichAxis:=", "X"), _
    Array("NAME:Attributes", "Name:=", "Port2", _
    "Flags:=", "", "Color:=", "(0 0 255)", _
    "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", _
    "MaterialValue:=", Chr(34) & "vacuum" & Chr(34), _
    "SolveInside:=", true)

'Assign wave ports
Set oModule = oDesign.GetModule("BoundarySetup")

'Port 1 assignment
oModule.AssignWavePort Array("NAME:Port1", _
    "Objects:=", Array("Port1"), _
    "NumModes:=", 1, _
    "RenormalizeAllTerminals:=", true, _
    "UseLineModeAlignment:=", false, _
    "DoDeembed:=", true, _
    "DeembedDist:=", "Lio", _
    Array("NAME:Modes", _
        Array("NAME:Mode1", _
            "ModeNum:=", 1, _
            "UseIntLine:=", true, _
            Array("NAME:IntLine", _
                "Start:=", Array(CStr(launch_points_left(0, 0)) & "mil", _
                    CStr(launch_points_left(0, 1)) & "mil", _
                    CStr(z_core2_bottom + layer_h_core2/2) & "mil"), _
                "End:=", Array(CStr(launch_points_left(0, 0)) & "mil", _
                    CStr(launch_points_left(0, 1)) & "mil", _
                    CStr(z_core2_bottom + cond_h) & "mil")), _
            "CharImp:=", "Zpi", _
            "AlignmentGroup:=", 0, _
            "RenormImp:=", "50ohm")), _
    "ShowReporterFilter:=", false, _
    "ReporterFilter:=", Array(true), _
    "UseAnalyticAlignment:=", false)

'Port 2 assignment
oModule.AssignWavePort Array("NAME:Port2", _
    "Objects:=", Array("Port2"), _
    "NumModes:=", 1, _
    "RenormalizeAllTerminals:=", true, _
    "UseLineModeAlignment:=", false, _
    "DoDeembed:=", true, _
    "DeembedDist:=", "Lio", _
    Array("NAME:Modes", _
        Array("NAME:Mode1", _
            "ModeNum:=", 1, _
            "UseIntLine:=", true, _
            Array("NAME:IntLine", _
                "Start:=", Array(CStr(launch_points_right(4, 0)) & "mil", _
                    CStr(launch_points_right(4, 1)) & "mil", _
                    CStr(z_core2_bottom + layer_h_core2/2) & "mil"), _
                "End:=", Array(CStr(launch_points_right(4, 0)) & "mil", _
                    CStr(launch_points_right(4, 1)) & "mil", _
                    CStr(z_core2_bottom + cond_h) & "mil")), _
            "CharImp:=", "Zpi", _
            "AlignmentGroup:=", 0, _
            "RenormImp:=", "50ohm")), _
    "ShowReporterFilter:=", false, _
    "ReporterFilter:=", Array(true), _
    "UseAnalyticAlignment:=", false)

'***************************************************************************
' Create Solution Setup
'***************************************************************************

Set oModule = oDesign.GetModule("AnalysisSetup")

'Main adaptive solution
oModule.InsertSetup "HfssDriven", Array("NAME:Setup1", _
    "Frequency:=", "CenterFreq", _
    "PortsOnly:=", false, _
    "MaxDeltaS:=", max_delta_s, _
    "UseMatrixConv:=", false, _
    "MaximumPasses:=", num_passes, _
    "MinimumPasses:=", 1, _
    "MinimumConvergedPasses:=", 1, _
    "PercentRefinement:=", percent_refine, _
    "IsEnabled:=", true, _
    "BasisOrder:=", 1, _
    "DoLambdaRefine:=", true, _
    "DoMaterialLambda:=", true, _
    "SetLambdaTarget:=", false, _
    "Target:=", 0.3333, _
    "UseMaxTetIncrease:=", false, _
    "PortAccuracy:=", 2, _
    "UseABCOnPort:=", false, _
    "SetPortMinMaxTri:=", false, _
    "UseDomains:=", false, _
    "UseIterativeSolver:=", false, _
    "SaveRadFieldsOnly:=", false, _
    "SaveAnyFields:=", true, _
    "NoAdditionalRefinementOnImport:=", false)

'Frequency sweep
oModule.InsertFrequencySweep "Setup1", Array("NAME:Sweep1", _
    "IsEnabled:=", true, _
    "RangeType:=", "LinearCount", _
    "RangeStart:=", start_freq, _
    "RangeEnd:=", end_freq, _
    "RangeCount:=", 1001, _
    "Type:=", sweep_type, _
    "SaveFields:=", false, _
    "SaveRadFields:=", false, _
    "InterpTolerance:=", 0.5, _
    "InterpMaxSolns:=", 250, _
    "InterpMinSolns:=", 0, _
    "InterpMinSubranges:=", 1, _
    "ExtrapToDC:=", false, _
    "InterpUseS:=", true, _
    "InterpUsePortImped:=", false, _
    "InterpUsePropConst:=", true, _
    "UseDerivativeConvergence:=", false, _
    "InterpDerivTolerance:=", 0.2, _
    "UseFullBasis:=", true, _
    "EnforcePassivity:=", true, _
    "PassivityErrorTolerance:=", 0.0001)

'***************************************************************************
' Create Reports
'***************************************************************************

Set oModule = oDesign.GetModule("ReportSetup")

'S-parameter magnitude plot
oModule.CreateReport "S_Parameters", "Modal Solution Data", "Rectangular Plot", _
    "Setup1 : Sweep1", Array("Domain:=", "Sweep"), _
    Array("Freq:=", Array("All")), _
    Array("X Component:=", "Freq", "Y Component:=", Array("dB(S(1,1))", "dB(S(2,1))")), _
    Array()

'Group delay plot
oModule.CreateReport "Group_Delay", "Modal Solution Data", "Rectangular Plot", _
    "Setup1 : Sweep1", Array("Domain:=", "Sweep"), _
    Array("Freq:=", Array("All")), _
    Array("X Component:=", "Freq", "Y Component:=", Array("GroupDelay(S(2,1))")), _
    Array()

'Phase plot
oModule.CreateReport "Phase_Response", "Modal Solution Data", "Rectangular Plot", _
    "Setup1 : Sweep1", Array("Domain:=", "Sweep"), _
    Array("Freq:=", Array("All")), _
    Array("X Component:=", "Freq", "Y Component:=", Array("ang_deg(S(2,1))")), _
    Array()

'***************************************************************************
' Validation and Save
'***************************************************************************

'Turn on material override
oDesign.SetDesignSettings Array("NAME:Design Settings Data", _
    "Allow Material Override:=", true)

'Validate design
oDesign.Validate

'Save project
oProject.Save

'Display completion message
MsgBox "8-Pole Multi-Layer Stripline Filter created successfully!" & vbCrLf & _
    "- 8 resonators (4 Type A, 4 Type B)" & vbCrLf & _
    "- Multi-layer stackup with Rogers 3010 and 2929 prepreg" & vbCrLf & _
    "- Complex launch structures" & vbCrLf & _
    "- No vias - pure stripline design" & vbCrLf & _
    "- Frequency: " & start_freq & " to " & end_freq
