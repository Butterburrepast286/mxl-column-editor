// Default managed form module for the MxlToHtml external data processor.
// LaunchParameter must contain a path to a JSON file with statusPath and either
// one inputPath/outputPath pair or an items array of such pairs.

&AtClient
Procedure OnOpen(Cancel)

	StatusPath = "";
	Completed = False;

	Try
		ConfigPath = TrimAll(LaunchParameter);

		If IsBlankString(ConfigPath) Then
			Raise "JSON launch configuration file is not specified";
		EndIf;

		Reader = New JSONReader;
		Reader.OpenFile(ConfigPath);
		RenderConfig = ReadJSON(Reader);
		Reader.Close();

		StatusPath = TrimAll(String(RenderConfig.statusPath));

		RenderItems = Undefined;
		If RenderConfig.Property("items", RenderItems) Then
			If RenderItems.Count() = 0 Then
				Raise "items array is empty";
			EndIf;
			For Each RenderItem In RenderItems Do
				RenderConfiguredItem(RenderItem);
			EndDo;
		Else
			RenderConfiguredItem(RenderConfig);
		EndIf;

		WriteStatus(StatusPath, True, "");
		Completed = True;

	Except
		ErrorText = ErrorDescription();
		Message(ErrorText);

		If Not IsBlankString(StatusPath) Then
			Try
				WriteStatus(StatusPath, False, ErrorText);
			Except
				Message("Failed to write status file: " + ErrorDescription());
			EndTry;
		EndIf;
	EndTry;

	If Completed Then
		Exit(False);
	EndIf;

EndProcedure

&AtClient
Procedure RenderConfiguredItem(RenderItem)

	InputFileName = TrimAll(String(RenderItem.inputPath));
	OutputFileName = TrimAll(String(RenderItem.outputPath));

	If IsBlankString(InputFileName) Then
		Raise "inputPath is not specified";
	EndIf;

	If IsBlankString(OutputFileName) Then
		Raise "outputPath is not specified";
	EndIf;

	ConvertMxlToHtml(InputFileName, OutputFileName);

EndProcedure

&AtServerNoContext
Procedure ConvertMxlToHtml(InputFileName, OutputFileName)

	SpreadsheetDocument = New SpreadsheetDocument;
	SpreadsheetDocument.Read(InputFileName);
	SpreadsheetDocument.Write(
		OutputFileName,
		SpreadsheetDocumentFileType.HTML);

EndProcedure

&AtClient
Procedure WriteStatus(StatusPath, Success, ErrorText)

	If IsBlankString(StatusPath) Then
		Return;
	EndIf;

	StatusData = New Structure;
	StatusData.Insert("success", Success);
	StatusData.Insert("error", ErrorText);

	Writer = New JSONWriter;
	Writer.OpenFile(StatusPath, "UTF-8");
	WriteJSON(Writer, StatusData);
	Writer.Close();

EndProcedure
