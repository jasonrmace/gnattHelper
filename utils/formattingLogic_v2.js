/* global Excel */

// ==============================================================================
// FORMATTING LOGIC ENGINE (Phase 13: The "Green Completion" Update)
// Replaces VBA 'GenerateSmartRules'
// 
// NEW FEATURE:
// - If % Complete (Col H) is 100%, the bar turns GREEN (#00B050).
// - This overrides the standard Grey progress bar.
//
// RETAINED FEATURES:
// - Global PTO Columns (Light Grey #EAEAEA)
// - Range Safety (Prevents crashes on infinite columns)
// - Cell Scrubbing (Clears manual colors before applying rules)
// ==============================================================================

window.FormattingLogic = {
    generateSmartRules: async (context) => {
        console.log("Formatting Engine: I AM LOCAL! (Phase 13 - Green 100%)");
        try {
            const sheet = context.workbook.worksheets.getItem("GanttChart");
            const teamSheet = context.workbook.worksheets.getItem("Team");

            // 1. DEFINE RANGES
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();
            
            const startRow = 7; // Row 8
            const endRow = footerRange.rowIndex - 1;
            const rowCount = endRow - startRow + 1;
            
            const headerRange = sheet.getRange("K6").getExtendedRange(Excel.KeyboardDirection.right);
            headerRange.load("columnCount");
            await context.sync();
            
            let colCount = headerRange.columnCount;
            if (colCount > 2000) colCount = 2000; 
            if (colCount < 1) colCount = 1;

            console.log(`Target Range: Rows=${rowCount}, Cols=${colCount} (Start: K8)`);

            // Define Ranges
            const gridRange = sheet.getRangeByIndexes(startRow, 10, rowCount, colCount); // Col K
            const namesRange = sheet.getRangeByIndexes(startRow, 2, rowCount, 1); // Col C

            // 2. CLEAR OLD RULES & FORMATS
            console.log("STEP 1: Clearing Rules & Scrubbing Colors...");
            
            gridRange.conditionalFormats.clearAll();
            namesRange.conditionalFormats.clearAll();
            
            // Scrub manual formatting to prevent "Ghost Colors"
            gridRange.format.fill.clear();
            namesRange.format.fill.clear();

            await context.sync(); 
            console.log(">> Success. Grid is clean.");

            // 3. LOAD TEAM COLORS
            console.log("STEP 2: Fetching Colors...");
            let teamRules = [];
            const teamTable = teamSheet.tables.getItem("Team");
            const nameCol = teamTable.columns.getItem("First Name").getDataBodyRange();
            const colorCol = teamTable.columns.getItem("Color").getDataBodyRange();
            
            nameCol.load("values");
            const colorProps = colorCol.getCellProperties({ format: { fill: { color: true } } });
            await context.sync();

            const names = nameCol.values;
            const colors = colorProps.value;

            for (let i = 0; i < names.length; i++) {
                const rawName = names[i][0];
                if (colors[i] && colors[i][0]) {
                    let hex = colors[i][0].format.fill.color;
                    
                    let isInvalid = !hex || hex === "null";
                    if (typeof hex === 'string') {
                        if (hex.toUpperCase() === "#FFFFFF") isInvalid = true;
                        if (hex.length !== 7 || hex[0] !== "#") isInvalid = true;
                    } else {
                        isInvalid = true;
                    }

                    if (rawName && !isInvalid) { 
                        teamRules.push({ 
                            name: rawName.toString().toUpperCase().trim(), 
                            color: hex 
                        });
                    }
                }
            }
            console.log(`>> Success. Loaded ${teamRules.length} Valid Team Rules.`);

            // =================================================================
            // APPLY RULES
            // =================================================================

            // --- BLOCK A0: 100% COMPLETE (GREEN OVERRIDE) ---
            // This MUST be added first to sit on top of the grey progress bar.
            console.log("STEP 3a: Applying 100% Green Completion...");
            const fComplete = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            // Logic: Inside Date Range AND % Complete (Col H) >= 1 (100%)
            fComplete.custom.rule.formula = '=AND(K$6>=$E8, K$6<=$F8, ISNUMBER($H8), $H8>=1)';
            fComplete.custom.format.fill.color = "#00B050"; // Green
            fComplete.stopIfTrue = false; // Allow borders to draw on top if needed
            await context.sync();

            // --- BLOCK A: STANDARD PROGRESS BAR (PARTIAL) ---
            console.log("STEP 3b: Applying Standard Progress Bar...");
            const fProg = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fProg.custom.rule.formula = '=AND(K$6>=$E8, K$6<=$F8, ISNUMBER($H8), ((K$6-$E8)/($F8-$E8+1)) < $H8)';
            fProg.custom.format.fill.color = "#5a5a5a"; // Dark Grey (User Choice)
            fProg.stopIfTrue = false;
            await context.sync(); 
            console.log(">> Progress Bars Applied.");

            // --- BLOCK B: TODAY BORDERS ---
            console.log("STEP 4: Applying Today Borders...");
            const fToday = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fToday.custom.rule.formula = '=K$6=TODAY()';
            
            fToday.custom.format.borders.getItem("EdgeLeft").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("EdgeLeft").color = "#FF0000";
            fToday.custom.format.borders.getItem("EdgeLeft").weight = Excel.BorderWeight.thick;
            
            fToday.custom.format.borders.getItem("EdgeRight").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("EdgeRight").color = "#FF0000";
            fToday.custom.format.borders.getItem("EdgeRight").weight = Excel.BorderWeight.thick;
            
            fToday.stopIfTrue = false;
            await context.sync();
            console.log(">> Today Borders Applied.");

            // --- BLOCK C: TEAM COLORS ---
            console.log("STEP 5: Applying Team Colors...");
            for (let member of teamRules) {
                const safeName = member.name.replace(/"/g, '""');
                
                // Name Cell
                const fName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fName.custom.rule.formula = `=UPPER(TRIM($C8))="${safeName}"`;
                fName.custom.format.fill.color = member.color;
                fName.stopIfTrue = true;

                // Grid Bar (Base Color)
                const fBar = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fBar.custom.rule.formula = `=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8, UPPER(TRIM($C8))="${safeName}")`;
                fBar.custom.format.fill.color = member.color;
                fBar.stopIfTrue = true; 
            }
            await context.sync(); 
            console.log(">> Team Colors Applied.");

            // --- BLOCK D: GENERIC BLUE ---
            console.log("STEP 6: Applying Generic Blue...");
            const fBlue = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fBlue.custom.rule.formula = '=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8)'; 
            fBlue.custom.format.fill.color = "#0070C0"; 
            fBlue.stopIfTrue = true; 
            await context.sync();
            console.log(">> Generic Blue Applied.");

            // --- BLOCK E: PARENT ROW ---
            console.log("STEP 7: Applying Parent Rows...");
            const fParent = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParent.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParent.custom.format.fill.color = "#D9D9D9"; 
            fParent.stopIfTrue = true; 
            
            const fParentName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParentName.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParentName.custom.format.fill.color = "#D9D9D9";
            fParentName.stopIfTrue = true;
            await context.sync();
            console.log(">> Parent Rows Applied.");

            // --- BLOCK F: PTO (GLOBAL COLUMN) ---
            console.log("STEP 8: Applying PTO (Global Column)...");
            const fPTO = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            
            // Check if Date (K6) is within Start/End for ANYONE in the list
            fPTO.custom.rule.formula = '=SUMPRODUCT((StartDate<=K$6)*((StartDate+NumberDays-1)>=K$6))>0';
            fPTO.custom.format.fill.color = "#EAEAEA"; 
            fPTO.stopIfTrue = false;
            await context.sync();
            console.log(">> PTO Applied.");

            // --- BLOCK G: HOLIDAYS ---
            console.log("STEP 9: Applying Holidays...");
            const fHol = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fHol.custom.rule.formula = '=COUNTIF(ListHolidays,K$6)>0';
            fHol.custom.format.fill.color = "#C8C8C8"; 
            fHol.stopIfTrue = false;
            await context.sync();
            console.log(">> Holidays Applied.");

            console.log("Formatting Engine: ALL RULES SUCCESSFUL.");

        } catch (error) {
            console.error("Formatting Logic Error at Step:", error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug Info:", error.debugInfo);
            }
        }
    }
};
