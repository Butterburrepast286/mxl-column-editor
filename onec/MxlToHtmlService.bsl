// Managed form module for MxlToHtmlService.epf -- the RESIDENT renderer.
//
// The original MxlToHtml.epf renders one document and exits, so every preview
// pays the platform's startup cost and flashes a window. This variant stays
// running and serves a queue instead.
//
// LaunchParameter is a SESSION FOLDER rather than a job file:
//
//     <session>\startup.log      every step this module reaches, appended
//     <session>\ready.json       written once the queue loop is armed
//     <session>\jobs\<id>.json   a job: {"items": [{name, inputPath, outputPath}]}
//     <session>\done\<id>.json   the outcome: {"success": ..., "error": ...}
//     <session>\stop             appears when the caller wants it to exit
//
// Everything is traced to startup.log, because a data processor launched with
// /Execute has nowhere useful to print: Message() goes to a window nobody is
// looking at. If startup.log is missing entirely, this module never ran at all
// -- which usually means the code was pasted somewhere other than the default
// managed form's module.
//
// Files are written with TextWriter rather than JSONWriter to keep the startup
// path as simple as possible; job files are still read with JSONReader.
//
// The conversion runs in server context because SpreadsheetDocument.Read is
// unavailable in the thin client.

&AtClient
Var SessionFolder;

&AtClient
Var JobsFolder;

&AtClient
Var DoneFolder;

&AtClient
Procedure OnOpen(Cancel)

	Try
		SessionFolder = TrimAll(LaunchParameter);

		If IsBlankString(SessionFolder) Then
			// Nowhere to report to; the caller will time out and say so.
			Return;
		EndIf;

		If Right(SessionFolder, 1) <> "\" Then
			SessionFolder = SessionFolder + "\";
		EndIf;

		Trace("OnOpen reached; session folder = " + SessionFolder);

		JobsFolder = SessionFolder + "jobs\";
		DoneFolder = SessionFolder + "done\";
		CreateDirectory(JobsFolder);
		CreateDirectory(DoneFolder);
		Trace("queue folders ready");

		WriteTextFile(SessionFolder + "ready.json", "{""ready"": true}");
		Trace("ready.json written");

		AttachIdleHandler("PollQueue", 0.3, True);
		Trace("idle handler attached; waiting for jobs");

	Except
		Trace("STARTUP FAILED: " + ErrorDescription());
	EndTry;

EndProcedure

&AtClient
Procedure PollQueue()

	Try
		StopFile = New File(SessionFolder + "stop");
		If StopFile.Exist() Then
			Trace("stop file seen; exiting");
			Exit(False);
			Return;
		EndIf;

		Jobs = FindFiles(JobsFolder, "*.json", False);
		For Each JobFile In Jobs Do
			ProcessJob(JobFile);
		EndDo;
	Except
		Trace("POLL FAILED: " + ErrorDescription());
	EndTry;

	// Re-arm: 1C only allows sub-second intervals for single-shot handlers, and
	// re-arming at the end also stops two passes from overlapping.
	AttachIdleHandler("PollQueue", 0.3, True);

EndProcedure

&AtClient
Procedure ProcessJob(JobFile)

	Success = True;
	ErrorText = "";
	Trace("job picked up: " + JobFile.Name);

	Try
		Reader = New JSONReader;
		Reader.OpenFile(JobFile.FullName);
		Job = ReadJSON(Reader);
		Reader.Close();

		RenderItems = Undefined;
		If Job.Property("items", RenderItems) Then
			If RenderItems.Count() = 0 Then
				Raise "items array is empty";
			EndIf;
			For Each RenderItem In RenderItems Do
				ConvertMxlToHtml(
					TrimAll(String(RenderItem.inputPath)),
					TrimAll(String(RenderItem.outputPath)));
			EndDo;
		Else
			ConvertMxlToHtml(
				TrimAll(String(Job.inputPath)),
				TrimAll(String(Job.outputPath)));
		EndIf;

	Except
		Success = False;
		ErrorText = ErrorDescription();
		Trace("job failed: " + ErrorText);
	EndTry;

	// Write the outcome before removing the job, so a crash between the two
	// cannot look like a job that silently vanished.
	Try
		Status = "{""success"": " + ?(Success, "true", "false")
			+ ", ""error"": """ + EscapedForJSON(ErrorText) + """}";
		WriteTextFile(DoneFolder + JobFile.Name, Status);
		Trace("job finished: " + JobFile.Name);
	Except
		Trace("could not write the status file: " + ErrorDescription());
	EndTry;

	Try
		DeleteFiles(JobFile.FullName);
	Except
		Trace("could not remove the job file: " + ErrorDescription());
	EndTry;

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
Function EscapedForJSON(Text)

	Result = StrReplace(Text, "\", "\\");
	Result = StrReplace(Result, """", "\""");
	Result = StrReplace(Result, Chars.CR, " ");
	Result = StrReplace(Result, Chars.LF, " ");
	Result = StrReplace(Result, Chars.Tab, " ");
	Return Result;

EndFunction

&AtClient
Procedure WriteTextFile(Path, Text)

	Writer = New TextWriter(Path, TextEncoding.UTF8);
	Writer.Write(Text);
	Writer.Close();

EndProcedure

&AtClient
Procedure Trace(Text)

	Try
		If IsBlankString(SessionFolder) Then
			Return;
		EndIf;
		Writer = New TextWriter(SessionFolder + "startup.log", TextEncoding.UTF8, , True);
		Writer.WriteLine(String(CurrentDate()) + "  " + Text);
		Writer.Close();
	Except
		// Tracing must never be the thing that breaks the renderer.
	EndTry;

EndProcedure
