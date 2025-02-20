


TEMPLATE = r"""
' -----------------------------------------------------------------------
' IPS Table Shell
'
' -----------------------------------------------------------------------
'Description: Run this script on the MDD prior to running the DMS script to export to Sav
'Loops through all of the Questions, and handles the assigning of Aliases and SubAliases at the VariableInstance Level.
'In certain instances, it grabs information from the parent field (such as with NumericGrids)

'--------------------------------------------------------------------------------------------------------------------------------------------
'Includes and Definitions

#include "Includes/DPrep/PrepMetadata.mrs"
#include "Includes/Globals.mrs"
#define SCRIPT_NAME	 "601_SavPrep.mrs"

'--------------------------------------------------------------------------------------------------------------------------------------------
'Setting up Dimmed variables and Constants for the script
Dim mDocument, mVariable, DataSrc, LoopCounter, LoopQuestion, savDscExists, x

savDscExists = false

Dim DataSource

'!
	Define data file to be converted to SPSS below using DATA_FILE_TO_USE.
	0 --> R data
	1 --> RM (merged) data,
	2 --> STK (stacked) data
!'
#Define DATA_FILE_TO_USE 0

'=======================================================
#if DATA_FILE_TO_USE == 0
	#Define S_MDD_SOURCE		PROCESSED_MDD
	#Define S_MDD 				SAV_Prep_MDD
	#define DATA_FILE 			Sav_DATA
#elif DATA_FILE_TO_USE == 1
	#Define S_MDD_SOURCE 		POSTMERGED_MDD
	#Define S_MDD 				SAV_Prep_MDD_MERGE
	#define DATA_FILE 			Sav_DATA_MERGE
#elif DATA_FILE_TO_USE == 2
	#Define S_MDD_SOURCE 		STK_MDD
	#Define S_MDD 				SAV_Prep_MDD_STK
	#define DATA_FILE 			Sav_DATA_STK
#endif
'=======================================================

'Limit space reserved for Other-Specify Text fields.
'	Regular text fields, Other-specs that already have max lengths, and non-text Other-specs will not be touched.
#Define TRUNCATE_OTHER_TEXT 	TRUE
#Define TRUNCATED_OTHER_SIZE	100

'-------------------------------------------------------------------------------------------------------------------------

'Delete an S-MDD if it finds one
Dim Sfso
Set Sfso = CreateObject("Scripting.FileSystemObject")

If (Sfso.FileExists(S_MDD)) Then
	Sfso.DeleteFile(S_MDD)
End If

'--------------------------------------------------------------------------------------------------------------------------------------------

'Open MDD to be transformed
Set mDocument = CreateObject("MDM.Document")
mDocument.Open(S_MDD_SOURCE, "", "&H0003")

'If locked, we'll need to add a new version in order to be able to create a new data source `
If mDocument.Versions.Latest.IsLocked Then
	mDocument.Versions.AddNew()
End If

'--------------------------------------------------------------------------------------------------------------------------------------------
'Remove any SAV Data Sources
For Each DataSource In mDocument.DataSources
	If LCase(DataSource.CDSCName) = LCase("mrSavDsc") Then
		savDscExists = True
		mDocument.DataSources.Remove(DataSource.Name)
	End If
Next

'Add new SAV Data Source base on Configurable Path
mDocument.DataSources.AddNew("mrSavDsc", "mrSavDsc", DATA_FILE)

'--------------------------------------------------------------------------------------------------------------------------------------------
'Setup the configuration properties of the generator

Dim MyConfiguration, MyGenerator
Dim oField, oCategory

For Each DataSource In mDocument.DataSources
	If LCase(DataSource.CDSCName) = LCase("mrSavDsc") Then
 		Set mDocument.DataSources.Current = DataSource

		Set MyGenerator = mDocument.AliasMap.Generator
		Set MyConfiguration = MyGenerator.Configuration
		MyConfiguration.MaxLength = 50

		' ProcessVariable(mDocument)
#include "{{{addin_file_path}}}"

		Set MyGenerator = Null
		Set MyConfiguration = Null
	End If
Next

'PRESERVE MORE THAN THE DEFAULT 50 CHARACTERS IN RESPONDENT.ID
if mDocument.Fields.Exist["Respondent"] then
    mDocument.Fields["Respondent"].Fields["ID"].MaxValue = 150
end if




Sub SetAlias(oField,Shortname)
    Dim oVariableInstance, ShouldSkip, iLevel, oIndexes, iIndexLooper, oParent, oIterationCat, sAlias
    Dim sVarMiddle
    For each oVariableInstance in oField.Variables
	    sVarMiddle = ""
        ShouldSkip = False
        iLevel = oVariableInstance.Indices.Count
        If iLevel>0 Then
            Set oParent = oVariableInstance
            oIndexes = Split(oVariableInstance.Indexes,",")
            If Find(oIndexes,"..") <> -1 Then
                ShouldSkip = True
            Else
                For iIndexLooper = iLevel to 1 Step -1
                    Set oParent = oParent.Parent.Parent
                    ' Check if Parent's Iterations are Categorical First...
                    If oIndexes[iIndexLooper - 1].Find("{") <> -1 Then
                        'Handle Classes in Loops (Go one more level Higher in parent hierarchy)
                        If CheckIsClassOrBlock(oParent) Then
                            Set oIterationCat = oParent.Parent.Categories[StripBraces(oIndexes[iIndexLooper - 1])]
                        Else
                            Set oIterationCat = oParent.Categories[StripBraces(oIndexes[iIndexLooper - 1])]
                        End If
                        sVarMiddle = ConvertToTripleDigitsText(oIterationCat.Properties["Value"]) + IIF(len(sVarMiddle) > 0, "_","") + sVarMiddle
                    Else
                        'Otherwise, just convert three digit Index
                        sVarMiddle = ConvertToTripleDigitsText(oIndexes[iIndexLooper - 1]) + IIF(len(sVarMiddle) > 0, "_","") + sVarMiddle
                    End If
                Next
            End If
        End If
        If Not ShouldSkip Then
            sAlias = SHortname + IIF(len(sVarMiddle) > 0, "_","") + sVarMiddle
            oVariableInstance.AliasName = sAlias
        End If
    Next
End Sub




''Top-Down Variable Processing
'Sub ProcessVariable(Q)
'	Dim oField, oSubField
'
'    set oField = Q
'    Select Case oField.ObjectTypeValue
'        Case 0 'Regular variable
'            'Check if Field needs to be processed at all
'            If oField.DataType=3 Then
'                SetupAnalysisValuesForSav(oField)
'            End If
'            SetupVariableForSAV(oField)
'        Case 1 'Array (Loop)						
'            'Process each Field and send alias prefix with it
'            SetupAnalysisValuesForSav(oField)
'            ProcessVariable(oField.Fields)
'        Case 2 'Grid						
'            'Process each Field and send alias prefix with it
'            SetupAnalysisValuesForSav(oField)
'            ProcessVariable(oField.Fields)
'        Case 3 'Class (Block, e.g. Bipolars)
'            ProcessVariable(oField.Fields)
'        Case 16
'            ' Skip
'        Case Else
'            Debug.Log("Unexpected field type found. If you are using an MDD straight from Excel Author, make sure to open and save in Professional before running this script. Found: " + oField.FullName + " - ObjectTypeValue = " + ctext(oField.ObjectTypeValue))
'            Debug.Echo("Unexpected field type found. If you are using an MDD straight from Excel Author, make sure to open and save in Professional before running this script. Found: " + oField.FullName + " - ObjectTypeValue = " + ctext(oField.ObjectTypeValue))
'            err.Raise(-999,"Unexpected field type found. If you are using an MDD straight from Excel Author, make sure to open and save in Professional before running this script. Found: " + oField.FullName + " - ObjectTypeValue = " + ctext(oField.ObjectTypeValue))
'    End Select
'    If oField.ObjectTypeValue=1 or oField.ObjectTypeValue=2 or oField.ObjectTypeValue=3 or oField.ObjectTypeValue=16 Then
'        For each oField in Q.Fields
'            ProcessVariable(Q)
'        Next
'    End If
'    if oField.ObjectTypeValue<>16 Then
'        For each oField in Q.HelperFields
'            ProcessVariable(Q)
'        Next
'    End If
'
'End Sub
'
'Sub SetupVariableForSAV(oField)
'    Dim oVariableInstance
'    For each oVariableInstance in oField.Variables
'        SetAlias(oVariableInstance)
'    Next
'End Sub
'
'Sub SetupAnalysisValuesForSav(oField)
'    Dim oVariableInstance
'    For each oVariableInstance in oField.Categories
'        SetAlias(oVariableInstance)
'    Next
'End Sub
'
'Sub SetAlias(oVariable)
'    Dim ccat
'    Select Case oVariable.FullName
'        Case "Timers.CustomSection[..].Seconds"
'            ccat = ccat
'        ' {{\{CODE_ADD}\}}
'        Case Else
'            debug.Log("Field not processed: "+ctext(oVariable.FullName)+"")
'            debug.Echo("Field not processed: "+ctext(oVariable.FullName)+"")
'            err.Raise(-999,"Field not processed: "+ctext(oVariable.FullName)+"")
'    End Select
'End Sub
'



Function ConvertToTripleDigitsText(iIteration)
	If iIteration < 10 Then
		ConvertToTripleDigitsText = "00" + ctext(iIteration)
	ElseIf iIteration < 100 Then
		ConvertToTripleDigitsText = "0" + ctext(iIteration)
	Else
		ConvertToTripleDigitsText = ctext(iIteration)
	End If
End Function




Function StripBraces(sString)
	StripBraces = Replace(Replace(sString,"{",""),"}","")
End Function





ErrorHandler:

If Err.Number <> 0 Then
	Debug.Log("An Error has occurred on line " + ctext(Err.LineNumber))
	Debug.Log("Description: " + Err.Description)
End If

'------------------------------------------------------------------------------------------------------------------------------------------------
'Save copy of transformed MDD to be used for the DMS export to Sav

mDocument.Save(S_MDD)
mDocument.Close()
'--------------------------------------------------------------------------------------------------------------------------------------------

"""