#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.

global FUNCTION_POINTERS := {}
global MODULE_PTR := 0
global OFF := 0
global ON := 1
OnExit("ExitScript")

Initialize() {
  VM_INSTALL_PATH := "C:\Program Files (x86)\VB\Voicemeeter"
  Dll_Name := "VoicemeeterRemote64.dll" 
  Module_Ptr := DllCall("LoadLibrary", "Str", VM_INSTALL_PATH . "\" . Dll_Name, "Ptr")
  return Module_Ptr
}

Get_Function_Pointers(Module_Ptr) {
  Function_Prefix := "VBVMR_"
  Function_Names := ["Login", "Logout", "RunVoicemeeter", "GetParameterFloat", "SetParameterFloat", "IsParametersDirty", "GetLevel", "GetMidiMessage", "Input_GetDeviceNumber", "Input_GetDeviceDescA", "Output_GetDeviceNumber", "Output_GetDeviceDescA"]
   
  function_pointers := {}
  functions := ""
  for index, function_name in Function_Names
  {
    full_function_name := Function_Prefix . function_name
    function_ptr := DllCall("GetProcAddress", "Ptr", Module_Ptr, "AStr", full_function_name, "Ptr")
    function_pointers[function_name] := function_ptr
    functions := functions . function_name . " : " . function_ptr . "`n"
  }
  ;MsgBox % functions
  return function_pointers
}

Login() {
  return DllCall(FUNCTION_POINTERS["Login"], "Int")
}

Logout() {
  return DllCall(FUNCTION_POINTERS["Logout"], "Int")
}

Unload() {
  DllCall("FreeLibrary", "Ptr", MODULE_PTR)
}

GetParameter(parameter_name) {
  DllCall(FUNCTION_POINTERS["IsParametersDirty"]) ; Force a refresh of values
  parameter_value := 0
  status := DllCall(FUNCTION_POINTERS["GetParameterFloat"], "AStr", parameter_name, "Ptr", &parameter_value, "Int")
  ;MsgBox % "status: " . status
  parameter_value := NumGet(parameter_value, 0, "Float")
  ;MsgBox % "value: " . parameter_value
  return parameter_value
}

SetParameter(parameter_name, parameter_value) {
  status := DllCall(FUNCTION_POINTERS["SetParameterFloat"], "AStr", parameter_name, "Float", parameter_value, "Int")
  ;MsgBox % "status: " . status  
}

ExitScript(exit_reason, exit_code) {
  Logout()
  Unload() 
  return 0
}

Main() {
  MODULE_PTR := Initialize()
  FUNCTION_POINTERS := Get_Function_Pointers(MODULE_PTR)
  Login()
}

Main()
MsgBox % "Loaded..."

GetParameterNameFromChannel(channel) {
  StringUpper, channel, channel
  StringLeft, name, channel, 1
  StringRight, number, channel, 1
  parameter_name := ""
  if (name = "A") {
    parameter_name := "Bus[" . number-1 . "].Mute"
  } else if (name = "S") {
    parameter_name := "Strip[" . number . "].Mute"
  }
  return parameter_name
}

MuteOn(channel) {
  parameter_name := GetParameterNameFromChannel(channel)
  SetParameter(parameter_name, ON)
}

MuteOff(channel) {
  parameter_name := GetParameterNameFromChannel(channel)
  SetParameter(parameter_name, OFF)
}

ToggleMute(channel) {
  parameter_name := GetParameterNameFromChannel(channel)
  current_value := GetParameter(parameter_name)
  if (current_value = OFF) {   
    MuteOn(channel)
  } else {
    MuteOff(channel)
  }
}

^1::
MuteOff("A1")
MuteOn("A2")
MuteOn("A3")
Gui, -caption
Gui, Add, Text, w80 cGreen, A1
Gui, Add, Text, x+10 cRed w80, A2
Gui, Add, Text, x+10 cRed w80, A3
Gui, Show, Center NA, "Mute Status"
Sleep 500
Gui, Destroy
return

^2::
MuteOn("A1")
MuteOff("A2")
MuteOn("A3")
Gui, -caption
Gui, Add, Text, w80 cRed, A1
Gui, Add, Text, x+10 cGreen w80, A2
Gui, Add, Text, x+10 cRed w80, A3
Gui, Show, Center NA, "Mute Status"
Sleep 500
Gui, Destroy
return

^3::
MuteOn("A1")
MuteOn("A2")
MuteOff("A3")
Gui, -caption
Gui, Add, Text, w80 cRed, A1
Gui, Add, Text, x+10 cRed w80, A2
Gui, Add, Text, x+10 cGreen w80, A3
Gui, Show, Center NA, "Mute Status"
Sleep 500
Gui, Destroy
return

;GetParameter("Strip[0].Mute")
;GetParameter("Bus[1].Mute")
;SetParameter("Strip[0].Mono", 1)
;ToggleMute("A1")
;ToggleMute("A2")
;Gui 1: Default
;Gui Color, Green
;Gui Font, cWhite s22 Normal, Verdana
;Gui Add, Text, x9 y9 h30 w400 Center BackgroundTrans, Header test
;Gui Add, Text, x11 y9 h30 w400 Center BackgroundTrans, Header test
;Gui Add, Text, x9 y11 h30 w400 Center BackgroundTrans, Header test
;Gui Add, Text, x11 y11 h30 w400 Center BackgroundTrans, Header test
;Gui Font, cBlack s22 Normal, Verdana
;Gui Add, Text, x10 y10 h30 w400 Center BackgroundTrans, Header test
;Gui Show, x140 y100 h100 w400, Checking the heading
; 2
Gui, Color, Fuchsia
Gui, Font, s20
Gui, Add, Text, cWhite, Some text here...
Gui, Color, Green
Gui, Font, s20
Gui, Add, Text, border 10 cWhite, Some text here...
Gui, Show
return
