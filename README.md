# DotNetUFTBridge

This will explain how to execute the Dot Net classes in UFT Test.

# How it works?

UFT has an ability to interact with dot net classes using DotNetFactory instance. We are utilizing the capabilities of DotNetFactory and Microsoft.CSharp.CSharpCodeProvider and System.CodeDom.Compiler assemblies to compile at run time and invoke the methods inside the dot net class.

# How do test the code?

1) Download the Bridge folder into C:\Git
2) Create a new blank UFT Test
3) Download the Test/Script.mts file and Replace the new contents.

```vbscript
'Load the Dot Net Bridge Core Class. Assume that you have deployed this code into C:\Git folder
LoadFunctionLibrary "C:\Git\Bridge\Compiler\DotNetBridge.vbs"
If IsObject(BridgeCore) Then
	BridgeCore.DotNetSource = "C:\Git\Bridge\C#\DotNetBridge.WindowActions.cs"
	blnIsWindowExists = BridgeCore.GetCompiledAssembly().HasWindowExists("Calculator")
	Msgbox blnIsWindowExists
End If
```

