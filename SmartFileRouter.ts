# Power Automate Desktop - Smart File Router
# Logic: Reads an Excel mapping table and moves PDFs based on multi-criteria string matching.

Excel.LaunchExcel.LaunchAndOpenUnderExistingProcess Path: $'''C:\Path\To\Your\Rules_Matrix.xlsx''' Visible: True ReadOnly: False UseMachineLocale: False Instance=> ExcelInstance
Excel.ReadFromExcel.ReadAllCells Instance: ExcelInstance GetCellContentsMode: Excel.GetCellContentsMode.TypedValues FirstLineIsHeader: True RangeValue=> RulesTable
Excel.CloseExcel.Close Instance: ExcelInstance

Folder.GetFiles Folder: $'''C:\Path\To\Source\Folder''' FileFilter: $'''*.pdf''' IncludeSubfolders: False FailOnAccessDenied: True SortBy1: Folder.SortBy.NoSort SortDescending1: False Files=> SourceFiles

SET TransferredFilesList TO $''''''

LOOP FOREACH CurrentFile IN SourceFiles
    SET match_found TO $'''False'''
    
    # Iterate through the Excel Rules Table
    LOOP FOREACH CurrentRule IN RulesTable
        IF IsNotEmpty(CurrentRule[0]) THEN
            # CRITERIA 1: File name must contain the Supplier Name (Column A)
            IF Contains(CurrentFile.NameWithoutExtension, CurrentRule[0], True) THEN
                # CRITERIA 2: File name must start with the Prefix (Column B)
                IF StartsWith(CurrentFile.NameWithoutExtension, CurrentRule[1], True) THEN
                    
                    # ACTION: Move file to destination path (Column C)
                    File.Move Files: CurrentFile Destination: CurrentRule[2] IfFileExists: File.IfExists.DoNothing MovedFiles=> MovedFiles
                    
                    # Log the move for the final report
                    SET TransferredFilesList TO $'''%TransferredFilesList%%CurrentFile.Name%  ->  %CurrentRule[2]%
'''
                    SET match_found TO $'''True'''
                END
            END
        END
    END
END

# Final User Reporting
IF IsEmpty(TransferredFilesList) THEN
    Display.ShowMessageDialog.ShowMessage Title: $'''Process Complete''' Message: $'''No new files were transferred.''' Icon: Display.Icon.Warning Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False ButtonPressed=> ButtonPressed
ELSE
    Display.ShowMessageDialog.ShowMessage Title: $'''Transfer Complete''' Message: $'''The following files were successfully transferred:

%TransferredFilesList%''' Icon: Display.Icon.Information Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False ButtonPressed=> ButtonPressed
END
