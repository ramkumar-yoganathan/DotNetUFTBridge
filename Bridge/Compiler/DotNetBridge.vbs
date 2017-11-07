Class DotNetBridge
    'Variable declaration
    Private objCodeProvider
    Private objCompilerParams
	Private strFlowMapFile
	Private objFlowMap
	Private strClassSource
	'Define the dot net source outside of the class. i.e from your test
	Public Property Let DotNetSource(ByVal ClassSource)
		strClassSource = ClassSource
	End Property
	'Get the flow map property file path for internal usage
	Public Property Get DotNetSource()
		DotNetSource = strClassSource
	End Property
    Private Sub Class_Initialize()
        Set objCodeProvider = GetFactory("Microsoft.CSharp.CSharpCodeProvider")
        Set objCompilerParams = GetFactory("System.CodeDom.Compiler.CompilerParameters")
    End Sub
	'''<summary>
	'''		Routine to get the instance of the DotNetFactory
	'''</summary>
	'''<param name="InstanceString" type="String">Name of the instance which you want create.</param>
	'''
	'''<returns type="Object">Factory Instance of Compiler. (In this case)</returns>
	''' 
	''' 
	'''<modification>
	'''		
	'''</modification>
	'''    
    Private Function GetFactory(ByVal InstanceString)
    	'Variable declaration
        Dim objDotNetFactory
        'Create factory instance
        Set objDotNetFactory = DotNetFactory.CreateInstance(InstanceString)
        'Check if it is valid else null
        If IsObject(objDotNetFactory) Then
            Set GetFactory = objDotNetFactory
        Else
        	'Something went wrong with UFT. Unable to create the DotNetFactory instance. 
            Set GetFactory = Null
        End If
    End Function    
	'''<summary>
	'''		Routine to set the compiler configuration. Include the third party assemblies, if in case
	'''		you are extending this class to further with the assembly.
	'''</summary>
	'''<param name="None" type="None"></param>
	'''
	'''<returns type="None">None</returns>
	'''<modification>
	'''		
	'''</modification>
	'''        
    Private Sub SetCompiler()
    	'Define Compiler Preference
        objCompilerParams.GenerateInMemory = True
        'Say No to generate executable, since you are in the UFT and Utilizing the assembly runtime.
        objCompilerParams.GenerateExecutable = False
        'Include the default assemblies of dotnet.
        objCompilerParams.ReferencedAssemblies.Add("System.dll")
        objCompilerParams.ReferencedAssemblies.Add("System.Windows.Forms.dll")
        'Include third party assembly.
        'CopyAssemblyForCompilation()
    End Sub
	'''<summary>
	'''		Routine to set the compiler configuration. Include the third party assemblies, if in case
	'''		you are extending this class to further with the assembly.
	'''</summary>
	'''<param name="" type=""></param>
	'''
	'''<returns type="String">Source code .</returns>
	''' 
	'''<modification>
	'''		
	'''</modification>
	'''            
    Private Function GetSourceCode()
    	'Variable declaration
    	Const FOR_READING = 1
        Set fileInstance = CreateObject("Scripting.FileSystemObject")
        'Open and Read the file
        Set fileReader = fileInstance.OpenTextFile(DotNetSource, FOR_READING)
        strSourceContent = fileReader.ReadAll()
        'Close the reader
        fileReader.Close()
        Set fileInstance = Nothing
        GetSourceCode = strSourceContent
    End Function
	'''<summar
	'''Routine to get the compiled assembly reference of the dot net action source code which was supplied 
	'''		
	'''</summary>
		'''
	'''<returns type="Object">Compiled assembly of the .</returns>
	''' 
	'''<modification>
	'''		
	'''</modification>
	'''
    Public Function GetCompiledAssembly()
    	'Variable declaration
        Dim objCompilerOutput
        'Get the dot net source of the action
        Dim strSourceCode : strSourceCode = GetSourceCode()
        'Set the compiler parameters
        SetCompiler()
		'Compile the assembly and generate the runtime
        Set objCompilerOutput = objCodeProvider.CompileAssemblyFromSource(objCompilerParams, strSourceCode)
        If IsObject(objCompilerOutput) Then
        	'If any compile time errors. Please check your flow source file.
            If objCompilerOutput.Errors.Count >= 1 Then
                Reporter.ReportEvent micFail, "Compilation failed. Source code has an error.", objCompilerOutput.Errors.Get_Item(0).ErrorText
                GetCompiledAssembly = False
                Exit Function
            End If
            'Get the assembly
            Set objAssembly = objCompilerOutput.CompiledAssembly
            'Get the assembly types
            Set objAssemblyTypes = objAssembly.GetTypes()
            Set objTypeEnums = objAssemblyTypes.GetEnumerator()
            While objTypeEnums.MoveNext()
                Set objType = objTypeEnums.Current
            Wend
            'Get the runtime reference of the assembly.
            Set objAssemblyInstance = objAssembly.CreateInstance(objType.FullName)
            'Return the assmbly reference to process further.
            If IsObject(objAssemblyInstance) Then
				Set GetCompiledAssembly = objAssemblyInstance
			Else
				'It looks like assembly is not generated. Try the dot net code with Visual Studio/Visual Studio Code.
				Set GetCompiledAssembly = Null
            End If
        End If
    End Function
    'Finally Block
	Private Sub Class_Terminate()
        Set objCodeProvider = Nothing
        Set objCompilerParams = Nothing
    End Sub
End Class


Public BridgeCore

Set BridgeCore = New DotNetBridge
