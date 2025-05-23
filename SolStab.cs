public class SolutionStability
{
    private static Application _app;

    private const int AreaLabelCol = 2;
    private const int AssayLabelCol = 6;
    private const int LabelClaimLabelCol = 10;
    private const int LabelClaimLabelValCol = 11;

    private const int DefaultNumSampleStandards = 1;
    private const int DefaultNumConditions = 2;
    private const string ConditionLabel = "Time point";
    private const int DefaultNumImpurities = 1;
    private const int DefaultImpurityCols = 4;
    private const int DefaultFormulaStartRow = 3;
    private const int DefaultFormulaStartCol = 3;
    private const int AssayDiffFormulaStartCol = 7;

    private const int PotencyDiffFormulaStartCol1 = 11;
    private const int PotencyDiffFormulaStartCol2 = 12;

    private const string TempDirectoryName = "ARD tempfiles";

    // Takes a worksheet and some metadata. Conditions list has been built by replicating items in conditions; copied "number of impurities" times.
    private static void UpdateWorksheet(Worksheet sheet, List<string> sampleLabels, List<string> conditions, int numImpurities, int impuritySampleNumber)
    {
        if (sheet == null)
            return;

        bool wasProtected = WorksheetUtilities.SetSheetProtection(sheet, null, false);

        // Expand or contract the data tables based on the number of conditions
        string[] conditionTables = new string[]
        {
            "ImpurityConditionsTable",
            "ResultsConditionsTable",
            "ArrayConditionsTable",
            "AssayConditionsTable",
            "ClaimsConditionsTable",
            "ImpuritySummaryConditionsTable"
        };

        ProcessMainLevel(sheet, conditions, conditionTables);
        ProcessImpurityLevel(sheet, conditions, numImpurities);

        ResultsAnalysis(sheet, sampleLabels, conditions);
        HandleImpuritySamples(sheet, sampleLabels, conditions, numImpurities, impuritySampleNumber);

        WorksheetUtilities.ActivateValidNamedRanges(sheet);
        RenderArrayNumberOfImpurities(sheet, numImpurities);
        ScrollToTopLeft(sheet);

        if (wasProtected)
            WorksheetUtilities.SetSheetProtection(sheet, null, true);
    }

    private static void ProcessMainLevel(Worksheet sheet, List<string> conditions, string[] nameConditionTables)
    {
        if (conditions.Count <= DefaultNumConditions)
            return;

        int numConditions = conditions.Count - DefaultNumConditions;

        WorksheetUtilities.InsertRows(sheet, "ImpurityConditionsTable", true, XlDirection.xlDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove, numConditions);
        WorksheetUtilities.InsertRows(sheet, "ResultsConditionsTable", true, XlDirection.xlDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove, numConditions);
        WorksheetUtilities.InsertRows(sheet, "ArrayConditionsTable", true, XlDirection.xlDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove, numConditions);
        WorksheetUtilities.InsertRows(sheet, "AssayConditionsTable", true, XlDirection.xlDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove, numConditions);
        WorksheetUtilities.InsertRows(sheet, "ClaimsConditionsTable", true, XlDirection.xlDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove, numConditions);
        WorksheetUtilities.InsertRows(sheet, "ImpuritySummaryConditionsTable", true, XlDirection.xlDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove, numConditions);

        WorksheetUtilities.SetNamedRangeValues(sheet, "ResultsConditions1", conditions);
        WorksheetUtilities.SetNamedRangeValues(sheet, "ImpurityConditions", conditions);
        WorksheetUtilities.SetNamedRangeValues(sheet, "ArrayConditions", conditions);
        WorksheetUtilities.SetNamedRangeValues(sheet, "AssayConditions", conditions);
        WorksheetUtilities.SetNamedRangeValues(sheet, "ClaimsConditions", conditions);
        WorksheetUtilities.SetNamedRangeValues(sheet, "ImpuritySummaryConditions", conditions);
    }

    private static void ProcessImpurityLevel(Worksheet sheet, List<string> conditions, int numImpurities)
    {
        for (int i = 1; i <= numImpurities; i++)
        {
            string impurityName = "Impurity" + i;
            string previousImpurityRangeName = "Impurity" + (i - 1);
            string impuritySummaryRangeName = "ImpuritySummary" + i;
            string diffRangeName = "ImpurityDifferenceCondition" + i;
            string previousDiffRangeName = "ImpurityDifferenceCondition" + (i - 1);
            string summaryConditionsRangeName = "ImpuritySummaryConditions" + i;
            string previousSummaryConditionsRangeName = "ImpuritySummaryConditions" + (i - 1);
            string diffPercentage = "ImpurityDifferencePercentage" + i;
            string previousDiffPercentage = "ImpurityDifferencePercentage" + (i - 1);

            if (!WorksheetUtilities.NamedRangeExists(sheet, impuritySummaryRangeName))
            {
                WorksheetUtilities.CopyNamedRangeToNewName(sheet, previousImpurityRangeName, impurityName, 1, DefaultImpurityCols + 1);
                WorksheetUtilities.CopyNamedRangeToNewName(sheet, previousSummaryConditionsRangeName, summaryConditionsRangeName, 1, 2);
                WorksheetUtilities.CopyNamedRangeToNewName(sheet, previousDiffRangeName, diffRangeName, 1, DefaultImpurityCols + 1);
                WorksheetUtilities.CopyNamedRangeToNewName(sheet, previousDiffPercentage, diffPercentage, 1, DefaultImpurityCols + 1);

                WorksheetUtilities.CopyNamedRangeToNewName(sheet, previousSummaryConditionsRangeName, summaryConditionsRangeName, 1, 2);
                WorksheetUtilities.LinkTwoNamedRangesCells(sheet, diffRangeName, summaryConditionsRangeName, 2, 2, false, true);
            }
        }
    }
    private static void HandleSamples(Worksheet sheet, IList<string> samplesStandards, IList<string> conditions)
    {
        // Handle the number of samples/standards
        int sampleCount = 0;
        int resultCount = 0;
        int arrayCount = 0;
        int sampleIndex = 0;

        foreach (string sample in samplesStandards)
        {
            sampleCount++;
            string label = "";
            string type = "";
            string name = "";
            string labelClassVal = "";

            string[] sampleStandardsArray = sample.Split(',');
            sampleStandardsArray = sampleStandardsArray.Select(s => s.Trim()).ToArray();

            if (sampleStandardsArray.Length >= 1)
                type = sampleStandardsArray[0].Trim();

            if (sampleStandardsArray.Length >= 2)
                labelClassVal = sampleStandardsArray[1].Trim();

            // Handle Results table
            if (!WorksheetUtilities.NamedRangeExists(sheet, "ResultsSampleStandards" + sampleCount))
            {
                // Add new named table
                // From 1 range between the numbers numbered 1 to 2, copy from range 1 to new range (copy all)
                WorksheetUtilities.InsertNamedRange(conditions.Count + 1, "Results", false, XlDirection.xlToRight, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRange(sheet, "ResultsSampleStandards" + (sampleCount - 1), "ResultsSampleStandards" + sampleCount);
                WorksheetUtilities.CopyNamedRange(sheet, "ResultsSampleRelinking" + (sampleCount - 1), "ResultsSampleRelinking" + sampleCount);
            }
        }
    }
    // Need a new NamedRange for handle the linking
    WorksheetUtilities.AddNamedRange(sheet, "ResultSampleLinking" + (sampleCount - 1), "ResultSampleLinking" + sampleCount,
        conditions.Count * 3 + 1, XlPasteType.xlPasteAll);
    // New val Report Table
    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "TableForAssay" + (sampleCount - 1), "TableForAssay" + sampleCount, conditions.Count * 3 + 1, XlPasteType.xlPasteAll);

    // set sample name
    WorksheetUtilities.SetNamedRangeValue(sheet, "ResultsSamplesStandards" + sampleCount, label, 1, 1);

    string rangeName = null;
    string linkingName = null;
    // Handle AreaSampleStandard Table
    if (!string.IsNullOrEmpty(type))
    {
        switch (type.ToUpper())
        {
            case "AREA":
                rangeName = "AreaSampleStandards" + (++areaCount);
                linkingName = "AreaSampleLinking" + (areaCount);
                header = "AreaClassification";
                if (!WorksheetUtilities.NamedRangeExists(sheet, rangeName))
                {
                    // add new named table
                    WorksheetUtilities.InsertNewNamedRange(conditions.Count + 2, sheet, "AreaData", false, XlDirection.xlUp,
                        XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "AreaSampleStandards" + (areaCount - 1), rangeName, conditions.Count * 3, 4,
                        XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "AreaSampleLinking" + (areaCount - 1), linkingName, conditions.Count * 3, 1,
                        XlPasteType.xlPasteAll);
                }

                // set sample name
                WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, label, 1, 1);
                break;

            case "ASSAY":
                rangeName = "AssaySampleStandards" + (++assayCount);
                    linkingName = "AssaySampleLinking" + (assayCount);
                    header = "AssaySampleHeader";
                    if (!WorksheetUtilities.NamedRangeExists(sheet, rangeName))
                    {
                        WorksheetUtilities.InsertRowsIntoNamedRange(conditions.Count + 2, sheet, "AssayData", false, XlDirection.xlUp,
                            XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "AssaySampleStandards" + (assayCount - 1), rangeName, conditions.Count * 3, 4,
                            XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "AssaySampleLinking" + (assayCount - 1), linkingName, conditions.Count * 3,
                            1, XlPasteType.xlPasteAll);
                    }

                    WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, label, 1, 1);
                    break;

                case "CLAIM":
                    rangeName = "ClaimSampleStandards" + (++claimCount);
                    linkingName = "ClaimSampleLinking" + (claimCount);
                    header = "ClaimSampleHeader";
                    if (!WorksheetUtilities.NamedRangeExists(sheet, rangeName))
                    {
                        WorksheetUtilities.InsertRowsIntoNamedRange(conditions.Count + 2, sheet, "ClaimData", false, XlDirection.xlUp,
                            XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ClaimSampleStandards" + (claimCount - 1), rangeName, conditions.Count * 3,
                            4, XlPasteType.xlPasteAll);
                        WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ClaimSampleLinking" + (claimCount - 1), linkingName, conditions.Count * 3,
                            1, XlPasteType.xlPasteAll);
                    }

                    WorksheetUtilities.SetNamedRangeValue(sheet, rangeName, label, 1, 1);
                    break;
            }

            // link Results cell to raw data and others
            WorksheetUtilities.LinkTwoNamedRangeCells(sheet, linkingName, "ResultSampleRelinking" + sampleCount, 1, 4, 1, 3, true, false);
            WorksheetUtilities.LinkTwoNamedRangeCells(sheet, linkingName, "ResultSampleLinking" + sampleCount, 1, 4, 1, 3, true, false);
            if (!string.IsNullOrEmpty(header))
            {
                WorksheetUtilities.LinkTwoNamedRangeCells(sheet, header, "ResultSampleStandards" + sampleCount, -1, 4, 1, 3, false, false);
                // New second column header depending on last column of source
                if (header == "LabelSimilarityHeader")
                {
                    WorksheetUtilities.LinkTwoNamedRangeCells(sheet, header, "ResultsSamplesStandards" + sampleCount, -1, 7, -1, 0, false, false);

                    // Add Timepoints if the values X-axis (same for Assay / Area)
                    WorksheetUtilities.LinkTwoNamedRangeCells(sheet, linkingName, "ResultSampleLinking" + sampleCount, 1, -1, 2, false, false);
                }
                else
                {
                    WorksheetUtilities.SetNamedRangeValue(sheet, "ResultsSamplesStandards" + sampleCount, "=IF(" + "\"AI\"" + "<>\"AI\", \"" + "\", \"" + "\")", 1, 2);
                    WorksheetUtilities.LinkTwoNamedRangeCells(sheet, linkingName, "ResultSampleLinking" + sampleCount, -1, 10, -1, 2, false, false);

                    // Add Timepoints if the columns exist (different column number for label claim)
                    // WorksheetUtilities.LinkTwoNamedRangeCells(sheet, linkingName, "ResultSampleLinking" + sampleCount, -1, 10, -1, 2, false, false);
                }
            }

            // Add linking for conditions on summary table
            WorksheetUtilities.LinkTwoNamedRangeCells(sheet, linkingName, "ResultSampleSummary" + sampleCount, -1, 1, -1, 1, false, false);
            /// end each sample loop

    // -----------------------------------------------------------------------

    private static void HandleImpuritySamples(Worksheet sheet, IList<string> sampleStandards, IList<string> conditions, int numImpurities, int impuritySampleNumber)
    {
        // Add Impurity Samples by newly added textbox.
        if (impuritySampleNumber > 1)
        {
            for (var x = 1; x <= impuritySampleNumber; x++)
            {
                //// Insert Template named table
                // add new named table
                // Insert a range between the ranges numbered 1 & 2, copy from range 1 to new range (copy all)
                WorksheetUtilities.InsertRowsIntoNamedRange(conditions.Count + 1, sheet, "ImpurityData", false, XlDirection.xlUp, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySampleStandards" + (x - 1), "ImpuritySampleStandards" + x, conditions.Count * 3, 1, XlPasteType.xlPasteAll);

                // Add conditions for linking to summary tables
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityConditions" + (x - 1), "ImpurityConditions" + x, conditions.Count * 3, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityDifferences" + (x - 1), "ImpurityDifferences" + x, conditions.Count * 3, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityDifferenceConditions" + (x - 1), "ImpurityDifferenceConditions" + x, conditions.Count * 3, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityInitials" + (x - 1), "ImpurityInitials" + x, conditions.Count * 3, 1, XlPasteType.xlPasteAll);

                WorksheetUtilities.InsertRowsIntoNamedRange(conditions.Count + 1, sheet, "ImpuritySummaryData", false, XlDirection.xlUp);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummarySample" + (x - 1), "ImpuritySummarySample" + x, conditions.Count * 3, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryDifference1" + (x - 1), "ImpuritySummaryDifference1" + x, conditions.Count * 3, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryDefinitions" + (x - 1), "ImpuritySummaryDefinitions" + x, conditions.Count * 2, 1, XlPasteType.xlPasteAll);
                WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryConditions1" + (x - 1), "ImpuritySummaryConditions1" + x, conditions.Count * 2, 1, XlPasteType.xlPasteAll);

            for (var i = 1; i <= numImpurities; i++)
            {
                if (i == x)
                {
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryDifference" + (i - 1), "ImpuritySummaryDifference" + i, 2, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityDifference" + (i - 1), "ImpurityDifference" + i, 1, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryConditions1" + (i - 1), "ImpuritySummaryConditions1" + i, 1, XlPasteType.xlPasteAll);
                }
            }
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryConditions" + (i - 1) + x, "ImpuritySummaryConditions" + i + x,
                        1, 1, XlPasteType.xlPasteAll);
                    WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityDifferenceConditions" + (i - 1), "ImpurityDifferenceConditions" + i + x,
                        1, 1, DefaultImpurityCols + 1, XlPasteType.xlPasteAll);

                    // WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpurityInitial" + (i - 1) + x, "ImpurityInitial" + i + x, 1,
                    // DefaultImpurityCols + 1, XlPasteType.xlPasteAll);

                    WorksheetUtilities.LinkTwoNamedRangeCells(sheet, "ImpuritySummaryDifference" + i + x, "ImpuritySummaryConditions" + i + x, -1, -1, 1,
                        1, false, true);
                    WorksheetUtilities.LinkTwoNamedRangeCells(sheet, "ImpurityDifferenceConditions" + i + x, "ImpuritySummaryConditions" + i + x, -1, -1,
                        1, 1, false, true);
                }
                else
                {
                    WorksheetUtilities.LinkTwoNamedRangeCells(sheet, "ImpurityDifference" + i + x, "ImpuritySummaryDifference" + i + x, -1, 1, -1, 1,
                        false, false);
                    WorksheetUtilities.LinkTwoNamedRangeCells(sheet, "ImpurityDifferenceConditions" + i + x, "ImpuritySummaryConditions" + i + x, -1, -1,
                        1, 1, false, true);
                }
            }

            // set sample name
            string label = "";
            string[,] sampleStandardsArray = new string[1, 1];

            if (sampleStandards.Count > 0 && !string.IsNullOrEmpty(sampleStandards[x - 1]))
            {
                sampleStandardsArray = JumpListExtensions.To2DArray(sampleStandards[x - 1].Split(','));
            }
            else
            {
                sampleStandardsArray[0,0] = "Sample" + x;
            }

            label = sampleStandardsArray[0,0].Trim();
            label = sampleStandardsArray[0].Trim();

            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpuritySampleStandards" + x, label, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpuritySummarySample" + x, label, 1, 1);

            // Link impuritySummaryDefinition with ImpuritySample
            WorksheetUtilities.LinkTwoNamedRangeCells(sheet, "ImpuritySummaryDefinition" + x, "ImpuritySummaryDefinitions" + x, -1, 1, -1, 1, false, false);

            // WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "TableForImpurity" + x, "TableForImpurity" + x, conditions.Count + 1, XlPasteType.xlPasteAll);
            WorksheetUtilities.ResizeNamedRange(sheet, "TableForImpurity" + x, conditions.Count + 1, 0);

            // OR - Create for new Resizing
            WorksheetUtilities.ResizeNamedRange(sheet, "ImpuritySummaryDefinition", conditions.Count + 1, 0);
            for (var i = 1; i < numImpurities; i++)
            {
                WorksheetUtilities.ResizeNamedRange(sheet, "ImpuritySummaryImpurityColumn" + i, conditions.Count - 1, 0);
            }
        }
        else
        {
            string label = "";
            string[] sampleStandardsArray = new string[1];

            if (sampleStandards.Count > 0 && !string.IsNullOrEmpty(sampleStandards[0]))
            {
                sampleStandardsArray = sampleStandards[0].Split(',');
            }
            else
            {
                sampleStandardsArray[0] = "Sample1";
            }

            label = sampleStandardsArray[0].Trim();

            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpuritySampleStandards", label, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpuritySummarySample", label, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpuritySummarySample1", label, 1, 1);
        }
        else if (impuritySampleNumber == 0)
        {
            // Added as 30/03 Comment:
            WorksheetUtilities.DeleteNamedRangeRows(sheet, "ToDeleteImpuritySections");
        }
        else
        {
            string label = "";
            string[] sampleStandardsArray = new string[] { "" };

            if (sampleStandards.Count > 0 && !string.IsNullOrEmpty(sampleStandards[0]))
            {
                sampleStandardsArray = sampleStandards[0].Split(',');
            }
            else
            {
                sampleStandardsArray[0] = "Sample1";
            }

            label = sampleStandardsArray[0].Trim();

            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpuritySampleStandards1", label, 1, 1);
            WorksheetUtilities.SetNamedRangeValue(sheet, "ImpuritySummarySample1", label, 1, 1);
        }

        // extend impurity range
        for (int i = 1; i <= numImpurities; i++)
        {
            string impurityRangeName = "Impurity" + i;

            if (!WorksheetUtilities.NamedRangeExists(sheet, impurityRangeName))
                WorksheetUtilities.ResizeNamedRange(sheet, impurityRangeName, (conditions.Count + 2) * (impuritySampleNumber + 1), 0);
        }
            WorksheetUtilities.ResizeNamedRange(sheet, impurityRangeName, (conditions.Count + 2) * (impuritySampleNumber + 1), 0);
        }
    }

    private static void HandleLargeNumberOfImpurities(Worksheet sheet, int numImpurities)
    {
        // 09-07 - Replication Code for more than 6 Impurities  Validation/Exporting
        if (numImpurities > 5)
        {
            int stepOrder = 2;
            int rowsToInsert = WorksheetUtilities.GetNamedRangeRowCount(sheet, "ImpuritySummaryReplication1");
            for (var i = 1; i <= numImpurities / 2; i++)
            {
                if (i % 3 == 0)
                {
                    stepOrder = stepOrder + 1;
                    Range namedRangeForInsertRow = WorksheetUtilities.GetNamedRange(sheet, "ImpuritySummaryData");
                    int indexToInsertConditions = namedRangeForInsertRow.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                    int indexToInsertRow = WorksheetUtilities.GetNamedRangeRowCount(sheet, "ImpuritySummaryData");
                    Range baseCell = (Range)namedRangeForInsertRow.Cells[indexToInsertRow, 1];
                    string baseAddress = baseCell.get_Address(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    // Regex to replace address letter and leave only column number
                    string pattern = "[A-Z]+";
                    string result = Regex.Replace(baseAddress, pattern, string.Empty);

                    int indexToInsertColumn = int.Parse(result);

                    // ImpuritySummaryData - First column w/ last row empty
                    WorksheetUtilities.InsertRowsFromRowToNumber(sheet, rowsToInsert, sheet, "ImpuritySummaryData", true, XlDirection.xlDown,
                        XlPasteType.xlPasteAll, indexToInsertColumn);
            // Copy first column into the new rows
            // We take first the formatting, and then the values, we cannot do it direct way as there are formula in the condition columns
            WorksheetUtilities.CopyNamedRangeToNewNamedRange(sheet, "ImpuritySummaryReplication1", "ImpuritySummaryReplication" + 
                stepOrder, XlPasteType.xlPasteFormats);
            WorksheetUtilities.LinkTwoNamedRangeCells(sheet, "ImpuritySummaryReplication1", "ImpuritySummaryReplication" + 
                stepOrder, -1, 1, -1, 1, false, false);

            // Get the xth and every following column and paste them in the new location
            int impNumber = 0;

            Range inputRange = null;
            Range conditionsRange = null;
            int colOffset = 0;

            for (var x = 1; x <= 5; x++)
            {
                // Search named range to move
                string rangeMove = WorksheetUtilities.GetNamedRange(sheet, "ImpuritySummaryImpurityColumn" + (impNumber + x));
                conditionsRange = WorksheetUtilities.GetNamedRange(sheet, "ImpuritySummaryReplication" + stepOrder);

                if (rangeMove == null)
                {
                    break;
                }
                else
                {
                    // Variables to get where to move the range
                    Range firstCell = conditionsRange.Cells[1, 1];

                    // Get values on integers
                    int rowOffset = firstCell.Row;
                    int colOffset = firstCell.Column + x;

                    WorksheetUtilities.MoveNamedRange(sheet, "ImpuritySummaryImpurityColumn" + (impNumber + x), rowOffset, colOffset);
                }
            }
            colToResize = colToResize + 1;

            // Final
            ReleaseComObject(firstCell);
            ReleaseComObject(impToMove);

            // Search next impurity
            // impnumber = impnumber + 1;

            // Resize original named range (final step)
            WorksheetUtilities.InsertNamedRange(sheet, "TableForImpurity", rowsToInsert, 0);
            if (colToResize > 0)
            {
                WorksheetUtilities.ResizeNamedRange(sheet, "TableForImpurity", 0, colToResize);
            }

            // Clean
            ReleaseComObject(conditionsRange);
            ReleaseComObject(impToMove);
            ReleaseComObject(namedRangeForInsertRow);
            ReleaseComObject(baseCell);
        }
    }

    private static void ScrollToTopRight(Worksheet sheet)
    {
        try
        {
            sheet.Application.Goto(sheet.Cells[1, 1], true);
        }
        catch
        {
            Logger.LogMessage("Scroll of sheet failed in SolutionStability.UpdateSolutionStabilitySheet", EventLogEntryType.Error);
        }
    }
        Logger.LogMessage("Scroll of sheet failed in SolutionStability.UpdateSolutionStabilitySheet", EventLogEntryType.Error);
    }

    /// helper functions below here

    /// <summary>
    /// Updates the difference formulas to use correct cell address(es)
    /// Copy formula from base cell down, only update current cell address
    /// </summary>
    /// <param name="sheet">The worksheet</param>
    /// <param name="namedRangeBaseName">The base named range to update</param>
    /// <param name="baseCellRow">The row number of the base cell (from which the base formula is gotten)</param>
    /// <param name="baseCellCol">The column number of the base cell (from which the base formula is gotten)</param>
    /// <param name="colOffset">The column ref set number from each cell to get reference address (from which the base formula is gotten)</param>
    private static void updateDifferenceFormulas(Worksheet sheet, string namedRangeBaseName, int baseCellRow, int baseCellCol, int colOffset)
    {
        if (sheet == null || string.IsNullOrEmpty(namedRangeBaseName))
        {
            return;
        }

        Name name = null;
        Range range = null;

        try
        {
            name = GetNamedRange(sheet, namedRangeBaseName);
            if (name == null)
            {
                return;
            }

            range = name.RefersToRange;
            if (range == null)
            {
                return;
            }
            UpdateFormulasInRange(range, baseCellRow, baseCellCol, colOffset);
        }
        finally
        {
            ReleaseComObject(name);
            ReleaseComObject(range);
        }
    }

    /// <summary>
    /// Updates the result formulas to use correct cell address(es)
    /// Copy formula from base cell down, only update current cell address
    /// </summary>
    /// <param name="sheet">The worksheet</param>
    /// <param name="namedRangeBaseName">The base named range to update</param>
    /// <param name="baseCellRow">The row number of the base cell (from which the base formula is gotten)</param>
    /// <param name="baseCellCol">The column number of the base cell (from which the base formula is gotten)</param>
    /// <param name="colOffset">The column offset set number from each cell to get reference address (from which the base formula is gotten)</param>
    /// <param name="staticCellRow">Static row to fix base ref</param>
    private static void UpdateResultFormulas(Worksheet sheet, string namedRangeBaseName, int baseCellRow, int baseCellCol, int colOffset, int staticCellRow)
    {
        if (sheet == null || string.IsNullOrEmpty(namedRangeBaseName))
        {
            return;
        }

        Name name = null;
        Range range = null;

        try
        {
            name = GetNamedRange(sheet, namedRangeBaseName);
            if (name == null)
            {
                return;
            }

            range = name.RefersToRange;
            if (range == null)
            {
                return;
            }
            range = name.RefersToRange;
            if (range == null)
            {
                return;
            }

            UpdateFormulasInRange(range, baseCellRow, baseCellCol, colOffset, staticCellRow, staticCellColumn);
        }
        finally
        {
            ReleaseComObject(name);
            ReleaseComObject(range);
        }
    }

    private static void UpdateFormulasInRange(Range range, int baseCellRow, int baseCellCol, int colOffset, int? staticCellRow = null, int? staticCellColumn = null)
    {
        Range baseCell = range.Cells[baseCellRow, baseCellCol] as Range;
        if (baseCell == null)
        {
            return;
        }

        string baseFormula = baseCell.FormulaLocal.ToString();
        string baseAddress = ((Range)baseCell.Cells[1, colOffset]).get_AddressLocal(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

        ReleaseComObject(baseCell);

        for (int j = 1; j < range.Rows.Count - 1; j++)
        {
            Range cell = range.Cells[baseCellRow + j, baseCellCol] as Range;

            if (cell != null)
            {
                if (staticCellRow.HasValue && staticCellColumn.HasValue)
                if (staticCellRow.HasValue && staticCellColumn.HasValue)
                {
                    Range pointCell = range.Cells[staticCellRow.Value, staticCellColumn.Value] as Range;
                    if (pointCell != null)
                    {
                        string oldFormulaAddress = ((Range)pointCell.Cells[1, colOffset]).get_AddressLocal(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                        string addressToReplace = ((Range)baseCell.Cells[1, colOffset]).get_AddressLocal(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                        cell.Formula = cell.Formula.Replace(addressToReplace, oldFormulaAddress);

                        ReleaseComObject(pointCell);
                    }
                }
                else
                {
                    string newCellAddress = ((Range)cell.Cells[1, colOffset]).get_AddressLocal(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    cell.Formula = baseFormula.Replace(baseAddress, newCellAddress);
                }

                ReleaseComObject(cell);
            }
        }
    }

    private static Name GetNamedRange(Worksheet sheet, string namedRangeBaseName)
    {
        try
        {
            return sheet.Names.Item(namedRangeBaseName, Type.Missing, Type.Missing) as Name;
        }
        catch
        {
            return null;
        }
    }

    private static void ReleaseComObject(object obj)
    {
        if (obj != null && Marshal.IsComObject(obj))
        {
            while (Marshal.ReleaseComObject(obj) > 0) { }
        }
    }
}
