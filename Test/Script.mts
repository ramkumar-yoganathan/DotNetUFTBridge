'Load the Dot Net Bridge Core Class. Assume that you have deployed this code into C:\Git folder
LoadFunctionLibrary "C:\Git\Bridge\Compiler\DotNetBridge.vbs"
If IsObject(BridgeCore) Then
	BridgeCore.DotNetSource = "C:\Git\Bridge\C#\DotNetBridge.WindowActions.cs"
	blnIsWindowExists = BridgeCore.GetCompiledAssembly().HasWindowExists("Calculator")
	Msgbox blnIsWindowExists
End If
