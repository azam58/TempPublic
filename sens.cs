namespace Something
{
    public class Sensitivity
    {
        private static Application _app;

        private const int DefaultNumReps = 6;
        private const int DefaultNumImpurities = 1;
        private const int MaxNumImpurities = 8;

        public static string UpdateImpSensitivitySheet(string sourcePath, int numReps, int numImpurities)
        {
            string returnPath = "";
            try
            {
                returnPath = UpdateImpSensitivitySheet2(sourcePath, numReps, numImpurities);
            }
            catch (Exception ex)
            {
                Logger.LogMessage("An error occurred in the call to ImpuritySensitivity.UpdateImpSensitivitySheet. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);

                try
                {
                    if (_app.Workbooks.Count > 0)
                    {
                        _app.Workbooks.Close();
                    }
                    _app = null;
                }
                catch
                {
                    Logger.LogMessage("An error occurred in the call to ImpuritySensitivity.UpdateImpSensitivitySheet. Application failed to close workbooks. Message and stack trace are:\r\n" + ex.Message + "\r\n" + ex.StackTrace, Level.Error);
                }
                finally
                {
                    WorksheetUtilities.ReleaseExcelApp();
                }
            }
            return returnPath;
        }
        private static string UpdateImpSensitivitySheet2(string sourcePath, int numReps, int numImpurities)
        {
            if (!File.Exists(sourcePath))
            {
                Logger.LogMessage("Error in call to ImpuritySensitivity.UpdateImpSensitivitySheet. Invalid source file path specified.", Level.Error);
                return "";
            }

            // Generate a random temp path to save new workbook
            string savePath = WorksheetUtilities.CopyWorkbook(sourcePath, TempDirectoryName, "Impurity Sensitivity Results.xls");
            if (String.IsNullOrEmpty(savePath)) return "";

            // Try to open the file
            _app = WorksheetUtilities.GetExcelApp();
            _app.Workbooks.Open(savePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Workbook book = _app.Workbooks[1];
            Worksheet sheet = book.Worksheets[1] as Worksheet;

            if (sheet != null)
            {

                bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

                // As of 02-23, always delete the hidden section of the spreadsheet.
                WorksheetUtilities.DeleteNamedRangeRows(sheet, "ToDelete");

                int offset = 0;
                if (numReps != DefaultNumReps)
                {
                    if (numReps > DefaultNumReps)
                    {
                        offset = numReps - DefaultNumReps;
                    }
                    else
                    {
                        offset = -(DefaultNumReps - numReps);
                    }
                }

                if (numReps > DefaultNumReps)
                {
                    int numRowsToInsert = numReps - DefaultNumReps;

                    // Insert the extra rows into the named ranges in the raw data and validation results tables - do this first as later an insert/copy is simpler
                    WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "SampleNumsRawData", false, XlDirection.xlDown, XlPasteType.xlPasteAll);
                    WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "SNResultsFormulas", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
                    // WorksheetUtilities.RefreshFormulasForNamedRange(sheet, "SNResultsFormulas", 1);
                    // WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "InjectNumsValidationResults", false, XlDirection.xlDown, XlPasteType.xlPasteAll);
                    //WorksheetUtilities.InsertRowsIntoNamedRange(numRowsToInsert, sheet, "ValidationImpurityResponses1", true, XlDirection.xlDown, XlPasteType.xlPasteAll);
                    //WorksheetUtilities.InsertColumnsIntoNamedRange(numRowsToInsert, sheet, "Data", XlDirection.xlRight);
                }
                else if (numReps < DefaultNumReps)
                {
                    // Default to the minimum
                    int numRowsToRemove DefaultNumReps - numReps;
                    if (DefaultNumReps - numRowsToRemove < 2) numRowsToRemove = DefaultNumReps - 2;
                    WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "SampleNumsRawData", xlDirection.xlDown);
                    WorksheetUtilities.DeleteRowsFromNamedRange(numRowsToRemove, sheet, "SampleNumsRawData2", XlDirection.xlDown);
                }

                // Re-number the sample numbers in the rawdata table
                if (numReps <= 1) numReps = 2;
                List<string> sampleNumbers = new List<string>(0);
                for (int i = 1; i < numReps; i++) sampleNumbers.Add(i.ToString());
                WorksheetUtilities.SetNamedRangeValues(sheet, "SampleNumsRawData", sampleNumbers);
                WorksheetUtilities.SetNamedRangeValues(sheet, "SampleNumsRawData2", sampleNumbers);
                // WorksheetUtilities.SetNamedRangeValues (sheet, "InjectNumsValidationResults", sampleNumbers);

                // Handle the number of impurities
                if (numImpurities > DefaultNumImpurities)
                {
                    int numImpuritiesToInsert = numImpurities - DefaultNumImpurities;
                    // if (numImpurities > MaxNumImpurities) WorksheetUtilities. InsertColumns IntoNamedRange (num Impurities - MaxNumImpurities, sheet, "Data", XIDirection. xlToRight);

                    for (int i = 1; i <= numImpuritiesToInsert; i++)
                    {
                        // Copy the named ranges as needed for each impurity
                        int namedRangeNum = i + 1;
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityResults" + 1, "ImpurityResults" + namedRangeNum, 1, 3, xlPasteType.xlPasteAll);
                        WorksheetUtilities.SetNamedRangeValue(sheet, "ImpurityResults" + namedRangeNum, "Impurity" + namedRangeNum, 2, 2);
                        //Added as 12-2022
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "signalToNoiseResults" + i, "signalToNoiseResults" + namedRangeNum, 1, 4, XlPasteType.xlPasteAll);
                        //worksheetUtilities.SetNamedRangeValue(sheet, "SignalToNoiseResults" + namedRangeNum, "S/N", 2, 2);
                        // Hidden as 12-2022
                        // worksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityCalcs" + i, "ImpurityCalcs" + namedRangeNum, 1, 3, XlPasteType.xlPasteAll);
                        // worksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ValidationImpurityResults" + i, "ValidationImpurityResults" + namedRangeNum, 1, 2, XlPasteType.xlPasteAll);
                    }
                    // Expand the results table - Hidden as 12-2022
                    //worksheetUtilities.InsertRowsIntoSingleRowedNamedRange(numImpuritiesToInsert, sheet, "ResultsTable", false, XlPasteType.xlPasteFormulas);

                    // Expand the validation results table - Hidden as 12-2022
                    //worksheetUtilities.InsertRowsIntoSingleRowedNamedRange(numImpuritiesToInsert, sheet, "ValidationResultsTable", false, XlPasteType.xlPasteFormulas);

                    // Expand the validation result detail table - Hidden as 12-2022
                    //worksheetUtilities.InsertRowsIntoSingleRowedNamedRange(numImpuritiesToInsert, sheet, "ValidationImpurityResultDetailTable", false, XlPasteType.xlPasteFormulas);

                    // Update the result table formulas
                    updateResultTableFormulas(sheet, numImpurities);

                    //Change Formulas for Validation Report
                    worksheetUtilities.UpdateSensitivityFormulas(sheet, numImpuritiesToInsert, offset);

                    // Update the validation result table - Hidden as 12-2022
                    //UpdateValidationResultTableFormulas(sheet, numImpurities);
                    // update the validation result detail table formula - Hidden as 12-2022
                    //UpdateValidationResultDetailTableFormulas(sheet, numImpurities, numReps);


                }
                else
                {
                    // Update the result table formulas
                    updateResultTableFormulas(sheet, numImpurities);

                    // Update the validation result table - Hidden as 12-2022
                    //UpdateValidationResultTableFormulas(sheet, numImpurities);
                    // update the validation result detail table formula
                    //UpdateValidationResultDetailTableFormulas(sheet, numImpurities, numReps);
                }

                try
                {
                    _app.Goto(sheet.Cells[1, 1], true);
                }
                catch
                {
                    Logger.LogMessage("Scroll of sheet failed in ImpuritySensitivity.UpdateImpSensitivitySheet!", Level.Error);
                }

                if (wasProtected) worksheetUtilities.SetSheetProtection(sheet, null, true);

                while (Marshal.ReleaseComObject(sheet) >= 0) { }
            }

            _app.Workbooks[1].Save();

            while (Marshal.ReleaseComObject(book) >= 0) { }
            _app.Workbooks.Close();

            //while (Marshal.ReleaseComObject(_app) >= 0) { }
            _app = null;
            worksheetUtilities.ReleaseExcelApp();

            return savePath;
        }

        private static void UpdateResultTableFormulas(_Worksheet sheet, int numImpurities)
        {
            if (sheet == null || numImpurities <= 0) return;

            const int srcTableColIndex = 2;
            const int calcQuantitationRowIndex = 5;

            Name resultTableName = null;
            Range resultsTableRange = null;
            object objImpurityResultsNamedRange = null;
            Name impurityResultsTableName = null;
            Range impurityResultsTableRange = null;
            //Added 12-2022
            object sNResultsNamedRange = null;
            Name sNResultsTableName = null;
            Range sNResultsTableRange = null;
            //
            object objImpurityCalcsNamedRange = null;
            Name impurityCalcs1TableName = null;
            Range impurityCalcs1TableRange = null;
            Range srcCell = null;
            Range destCell = null;
            Range resultsTableRow = null;

            // Get the range by name - Hidden as 12/2022
            //object objResultsTableNamedRange = sheet.Names.Item("ResultsTable", Type.Missing, Type.Missing);
            //if (!(objResultsTableNamedRange is Name)) goto Cleanup_UpdateResultsTableFormulas;

            //resultTableName = objResultsTableNamedRange as Name;
            //resultsTableRange = resultTableName.RefersToRange;

            for (int i = 1; i <= numImpurities; i++)
            {

                objImpurityResultsNamedRange = sheet.Names.Item("ImpurityResults" + i, Type.Missing, Type.Missing);
                if (!(objImpurityResultsNamedRange is Name)) continue;

                //resultsTableRow = resultsTableRange.Rows[i, Type.Missing] as Range;

                impurityResultsTableName = objImpurityResultsNamedRange as Name;
                impurityResultsTableRange = impurityResultsTableName.RefersToRange;

                // Set the first cell of the current row - Impurity
                srcCell = impurityResultsTableRange.Cells[2, srcTableColIndex] as Range;
                if (srcCell != null)
                {
                    string srcCellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    srcCellAddress = srcCellAddress.Replace("$", "");

                    //destCell = resultsTableRow.Cells[1, 1] as Range;
                    //if (destCell != null)
                    //{
                    //    string destCellAddress = destCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    //    destCell.Value2 = "=" + srcCellAddress;
                    //while (Marshal.ReleaseComObject(destCell) >= 0) { }
                    //}
                    while (Marshal.ReleaseComObject(srcCell) >= 0) { }
                }

                // Set the third cell of the current row - % RSD at RL
                srcCell = impurityResultsTableRange.Cells[impurityResultsTableRange.Rows.Count, srcTableColIndex] as Range;
                if (srcCell != null)
                {
                    string srcCellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    srcCellAddress = srcCellAddress.Replace("$", "");

                    //destCell = resultsTableRow.Cells[1, 3] as Range;
                    //if (destCell != null)
                    //{
                    //    //string destCellAddress = destCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    //    srcCellAddress = String.Format("=IF({0}=\"\",\"\",{0})", srcCellAddress);
                    //    destCell.Value2 = srcCellAddress;
                    //while (Marshal.ReleaseComObject(destCell) >= 0) { }
                    //}
                    while (Marshal.ReleaseComObject(srcCell) >= 0) { }
                }

                //Hidden as 12/2022
                // Set the second cell of the current row - Calculated QL
                #region ImpurityCalcs
                //objImpurityCalcsNamedRange = sheet.Names.Item("ImpurityCalcs" + i, Type.Missing, Type.Missing);
                //if (!(objImpurityCalcsNamedRange is Name)) goto Cleanup_UpdateResultsTableFormulas;

                //impurityCalcs1TableName = objImpurityCalcsNamedRange as Name;
                //impurityCalcs1TableRange = impurityCalcs1TableName.RefersToRange;

                //srcCell = impurityCalcs1TableRange.Cells[calcQuantitationRowIndex, srcTableColIndex-1] as Range;
                //if (srcCell != null && resultsTableRow != null)
                //{
                //    string srcCellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                //    srcCellAddress = srcCellAddress.Replace("$", "");

                //    destCell = resultsTableRow.Cells[1, 2] as Range;
                //    if (destCell != null)
                //    {
                //        //string destCellAddress = destCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                //        //destCell.Value2 = String.Format("=IF({0}=\"\",\"\",{0})", srcCellAddress);
                //        while (Marshal.ReleaseComObject(destCell) >= 0) { }
                //    }
                //    while (Marshal.ReleaseComObject(destCell) >= 0) { }
                //}
                #endregion


                // Clean up
                try
                {
                    if (resultsTableRow != null) while (Marshal.ReleaseComObject(resultsTableRow) >= 0) { }
                    if (impurityResultsTableRange != null) while (Marshal.ReleaseComObject(impurityResultsTableRange) >= 0) { }
                    if (impurityResultsTableName != null) while (Marshal.ReleaseComObject(impurityResultsTableName) >= 0) { }
                    if (objImpurityResultsNamedRange != null) while (Marshal.ReleaseComObject(objImpurityResultsNamedRange) >= 0) { }
                    if (impurityCalcs1TableRange != null) while (Marshal.ReleaseComObject(impurityCalcs1TableRange) >= 0) { }
                    if (impurityCalcs1TableName != null) while (Marshal.ReleaseComObject(impurityCalcs1TableName) >= 0) { }
                    if (objImpurityCalcsNamedRange != null) while (Marshal.ReleaseComObject(objImpurityCalcsNamedRange) >= 0) { }
                }
                catch
                {
                    continue;
                }
            }

            //Added 12-2022 - SignalToNoise
            for (int i = 1; i <= numImpurities; i++)
            {
                sNResultsNamedRange = sheet.Names.Item("SignalToNoiseResults" + i, Type.Missing, Type.Missing);
                if (!(sNResultsNamedRange is Name)) continue;

                //resultsTableRow = resultsTableRange.Rows[i, Type.Missing] as Range;

                sNResultsTableName = sNResultsNamedRange as Name;
                sNResultsTableRange = sNResultsTableName.RefersToRange;

                // Set the first cell of the current row - S/N
                srcCell1 = sNResultsTableRange.Cells[1, srcTableColIndex] as Range;
                if (srcCell1 != null && resultsTableRow != null)
                {
                    string srcCellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    srcCellAddress = srcCellAddress.Replace("$", "");

                    //destCell = resultsTableRow.Cells[1, 1] as Range;
                    //if (destCell != null)
                    //{
                    //    //string destCellAddress = destCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    //    destCell.Value2 = "=" + srcCellAddress;
                    //    while (Marshal.ReleaseComObject(destCell) >= 0) { }
                    //}
                    while (Marshal.ReleaseComObject(srcCell) >= 0) { }
                }

                // Set the third cell of the current row - % RSD at RL
                srcCell = sNResultsTableRange.Cells[sNResultsTableRange.Rows.Count, srcTableColIndex] as Range;
                if (srcCell != null && resultsTableRow != null)
                {
                    string srcCellAddress = srcCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    srcCellAddress = srcCellAddress.Replace("$", "");

                    //destCell = resultsTableRow.Cells[1, 3] as Range;
                    //if (destCell != null)
                    //{
                    //    //string destCellAddress = destCell.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    //    srcCellAddress = String.Format("=IF({0}=\"\",\"\",{0})", srcCellAddress);
                    //    destCell.Value2 = srcCellAddress;
                    //while (Marshal.ReleaseComObject(destCell) >= 0) { }
                    //}
                    while (Marshal.ReleaseComObject(srcCell) >= 0) { }
                }

                // Clean up
                try
                {
                    if (resultsTableRow != null) while (Marshal.ReleaseComObject(resultsTableRow) >= 0) { }
                    if (sNResultsNamedRange != null) while (Marshal.ReleaseComObject(sNResultsNamedRange) >= 0) { }
                    if (sNResultsTableName != null) while (Marshal.ReleaseComObject(sNResultsTableName) >= 0) { }
                    if (sNResultsTableRange != null) while (Marshal.ReleaseComObject(sNResultsTableRange) >= 0) { }
                }
                catch
                {
                    continue;
                }
            }

        Cleanup_UpdateResultsTableFormulas:
            {
                try
                {
                    if (srcCell != null) while (Marshal.ReleaseComObject(srcCell) >= 0) { }
                    if (destCell != null) while (Marshal.ReleaseComObject(destCell) >= 0) { }
                    if (resultsTableRow != null) while (Marshal.ReleaseComObject(resultsTableRow) >= 0) { }
                    if (resultsTableRange != null) while (Marshal.ReleaseComObject(resultsTableRange) >= 0) { }
                    if (resultTableName != null) while (Marshal.ReleaseComObject(resultTableName) >= 0) { }
                    //if (objResultsTableNamedRange != null) while (Marshal.ReleaseComObject(objResultsTableNamedRange) >= 0) { }
                    if (impurityResultsTableRange != null) while (Marshal.ReleaseComObject(impurityResultsTableRange) >= 0) { }
                    if (impurityResultsTableName != null) while (Marshal.ReleaseComObject(impurityResultsTableName) >= 0) { }
                    if (objImpurityResultsNamedRange != null) while (Marshal.ReleaseComObject(objImpurityResultsNamedRange) >= 0) { }
                    if (impurityCalcs1TableRange != null) while (Marshal.ReleaseComObject(impurityCalcs1TableRange) >= 0) { }
                    if (impurityCalcs1TableName != null) while (Marshal.ReleaseComObject(impurityCalcs1TableName) >= 0) { }
                    if (objImpurityCalcsNamedRange != null) while (Marshal.ReleaseComObject(objImpurityCalcsNamedRange) >= 0) { }
                    if (sNResultsNamedRange != null) while (Marshal.ReleaseComObject(sNResultsNamedRange) >= 0) { }
                    if (sNResultsTableName != null) while (Marshal.ReleaseComObject(sNResultsTableName) >= 0) { }
                    if (sNResultsTableRange != null) while (Marshal.ReleaseComObject(sNResultsTableRange) >= 0) { }

                    // ReSharper disable RedundantAssignment
                    sheet = null;
                    // ReSharper restore RedundantAssignment

                }
                catch
                {
                    return;
                }
            }
        }
    }// end class
}
