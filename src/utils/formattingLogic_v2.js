/* global Excel */

// ==============================================================================
// FORMATTING LOGIC ENGINE (Phase 17: Project Row Protection)
// Replaces VBA 'GenerateSmartRules'
//
// CHANGES:
// 1. Block C (Team Colors): Added 'ISNUMBER(SEARCH(".", $A8))' to BOTH rules.
//    - Name Cell (Col C): Only colors if ID has a dot (Task).
//    - Grid Bar (Col K): Only colors if ID has a dot (Task).
//    - Project Rows (Whole Numbers): Ignored by Block C, caught by Block E (Grey).
// ==============================================================================

export const FormattingLogic = {
    /**
     * Performs a full reset of the sheet formatting.
     * Discovery logic for start/end rows stays here.
     */
    generateSmartRules: async (context, sheetName = "Houston") => {
        console.log("Formatting Engine: Starting...");
        
        // 1. SHOW SPINNER (Referencing global loader from index.jsx)
        if (window.GlobalLoader) window.GlobalLoader.show("Applying Rules...");

        try {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const teamSheet = context.workbook.worksheets.getItem("Team");

            // 1. DEFINE RANGES
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load(["rowIndex", "isNullObject"]);
            await context.sync();

            if (footerRange.isNullObject) {
                console.error("Formatting Engine: 'DO NOT DELETE' footer not found. Aborting.");
                return;
            }
            
            const startRow = 7; // Row 8
            const endRow = footerRange.rowIndex - 1;
            const actualRowCount = Math.max(endRow - startRow + 1, 1);
            
            const headerRange = sheet.getRange("K6:ZZ6"); 
            headerRange.load("columnCount");
            await context.sync();
            
            let colCount = headerRange.columnCount;
            if (colCount > 100) colCount = 100; 
            if (colCount < 10) colCount = 10;

            console.log(`Target Range: Rows=${actualRowCount}, Cols=${colCount} (Start: K8)`);

            // 2. CLEAR OLD RULES & FORMATS
            console.log("STEP 1: Clearing Rules & Scrubbing Colors...");
            const gridRange = sheet.getRangeByIndexes(startRow, 10, actualRowCount, colCount);
            const namesRange = sheet.getRangeByIndexes(startRow, 2, actualRowCount, 1);
            
            // Clear all before full re-apply
            gridRange.conditionalFormats.clearAll();
            namesRange.conditionalFormats.clearAll();
            gridRange.format.fill.clear();
            namesRange.format.fill.clear();

            // CRITICAL: Sync after clearing to reset the CF engine state
            await context.sync();

            await FormattingLogic.applyRulesToRange(context, sheetName, startRow, actualRowCount, colCount);
        } catch (error) {
            console.error("Formatting Logic Error:", error);
        } finally {
            if (window.GlobalLoader) window.GlobalLoader.hide();
        }
    },

    /**
     * Applies logic to a SPECIFIC range without clearing existing rules elsewhere.
     * Use this for new rows to ensure they inherit the Gantt behavior.
     */
    applyRulesToRange: async (context, sheetName, startRow, rowCount, colCount) => {
        try {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const teamSheet = context.workbook.worksheets.getItem("Team");
            
            const gridRange = sheet.getRangeByIndexes(startRow, 10, rowCount, colCount);
            const namesRange = sheet.getRangeByIndexes(startRow, 2, rowCount, 1);
            const r = startRow + 1;

            if (rowCount <= 0) return;

            // CRITICAL: Unmerge before applying CF rules. 
            // The Excel JS API often throws 'ItemNotFound' when adding CFs to ranges with partial merges.
            gridRange.unmerge();
            namesRange.unmerge();

            gridRange.format.fill.clear();
            namesRange.format.fill.clear();

            // LOAD TEAM COLORS
            console.log("STEP 2: Fetching Colors...");
            let teamRules = [];
            let officeMap = {}; // Map: { NAME: OFFICE }
            const teamTable = teamSheet.tables.getItemOrNullObject("Team");
            teamTable.load("isNullObject");
            await context.sync();

            if (teamTable.isNullObject) {
                console.warn("Formatting Engine: 'Team' table not found. Skipping team colors.");
            } else {
                const rowCount = teamTable.rows.getCount();
                await context.sync();

                if (rowCount.value > 0) {
                    const nameCol = teamTable.columns.getItem("First Name").getDataBodyRange();
                    const colorCol = teamTable.columns.getItem("Color").getDataBodyRange();
                    const officeCol = teamTable.columns.getItem("Office").getDataBodyRange();
                    nameCol.load("values");
                    officeCol.load("values");
                    const colorProps = colorCol.getCellProperties({ format: { fill: { color: true } } });
                    await context.sync();

                    const names = nameCol.values;
                    const colors = colorProps.value;
                    const offices = officeCol.values;
            
                    for (let i = 0; i < names.length; i++) {
                        const rawName = names[i][0];
                        const office = offices[i][0];
                        
                        if (rawName) {
                            officeMap[rawName.toString().toUpperCase().trim()] = office;
                        }

                        if (colors[i] && colors[i][0]) {
                            let hex = colors[i][0].format.fill.color;
                            let isInvalid = !hex || hex === "null" || (typeof hex === 'string' && (hex.toUpperCase() === "#FFFFFF" || hex.length !== 7));
                            
                            if (rawName && !isInvalid) { 
                                teamRules.push({ 
                                    name: rawName.toString().toUpperCase().trim(), 
                                    color: hex 
                                });
                            }
                        }
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
            fComplete.custom.rule.formula = `=AND(K$6>=$E${r}, K$6<=$F${r}, ISNUMBER($H${r}), $H${r}>=1)`;
            fComplete.custom.format.fill.color = "#00B050"; // Green
            fComplete.stopIfTrue = false; // Allow borders to draw on top if needed
            await context.sync();

            // --- BLOCK A: STANDARD PROGRESS BAR (PARTIAL) ---
            console.log("STEP 3b: Applying Standard Progress Bar...");
            const fProg = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fProg.custom.rule.formula = `=AND(K$6>=$E${r}, K$6<=$F${r}, ISNUMBER($H${r}), ((K$6-$E${r})/($F${r}-$E${r}+1)) < $H${r})`;
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

            // --- BLOCK C: TEAM COLORS (PROJECT PROTECTED) ---
            console.log("STEP 5: Applying Team Colors...");
            if (teamRules.length > 0) { // Only proceed if there are team rules to apply
                for (let i = 0; i < teamRules.length; i++) {
                    const member = teamRules[i];
                    const safeName = member.name.replace(/"/g, '""');
                    
                    // 1. NAME CELL (Strict Task Only)
                    const fName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                    fName.custom.rule.formula = `=AND(UPPER(TRIM($C${r}))="${safeName}", ISNUMBER(SEARCH(".", $A${r})))`;
                    fName.custom.format.fill.color = member.color;
                    fName.stopIfTrue = true;

                    // 2. GRID BAR (Strict Task Only)
                    const fBar = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                    fBar.custom.rule.formula = `=AND(ISNUMBER($E${r}), K$6>=$E${r}, K$6<=$F${r}, UPPER(TRIM($C${r}))="${safeName}", ISNUMBER(SEARCH(".", $A${r})))`;
                    fBar.custom.format.fill.color = member.color;
                    fBar.stopIfTrue = true; 
                    
                    // Intermediate sync every 3 members to flush the API command batch.
                    // Prevents ItemNotFound/exhaustion errors on complex range formatting.
                    if (i % 3 === 0 && i > 0) await context.sync();
                }
                await context.sync(); // Final sync after the loop
            } else {
                console.log(">> No Team Rules to Apply.");
            }
            console.log(">> Team Colors Applied.");

            // --- BLOCK D: GENERIC BLUE ---
            console.log("STEP 6: Applying Generic Blue...");
            const fBlue = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fBlue.custom.rule.formula = `=AND(ISNUMBER($E${r}), K$6>=$E${r}, K$6<=$F${r})`; 
            fBlue.custom.format.fill.color = "#0070C0"; 
            fBlue.stopIfTrue = true; 
            await context.sync();
            console.log(">> Generic Blue Applied.");

            // --- BLOCK E: PARENT ROW ---
            console.log("STEP 7: Applying Parent Rows...");
            const fParent = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParent.custom.rule.formula = `=AND($A${r}<>"", ISERROR(SEARCH(".", $A${r})))`;
            fParent.custom.format.fill.color = "#D9D9D9"; 
            fParent.stopIfTrue = true; 
            
            const fParentName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParentName.custom.rule.formula = `=AND($A${r}<>"", ISERROR(SEARCH(".", $A${r})))`;
            fParentName.custom.format.fill.color = "#D9D9D9";
            fParentName.stopIfTrue = true;
            await context.sync();
            console.log(">> Parent Rows Applied.");

            // --- BLOCK G: HOLIDAYS (SWAPPED UP) ---
            console.log("STEP 8: Applying Holidays...");
            // Using IFERROR/ISERROR safety in case ListHolidays named range doesn't exist yet
            const fHol = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fHol.custom.rule.formula = '=IFERROR(COUNTIF(ListHolidays,K$6)>0, FALSE)';
            fHol.custom.format.fill.color = "#C8C8C8"; 
            fHol.stopIfTrue = false;
            await context.sync();
            console.log(">> Holidays Applied.");

            // --- BLOCK F: PTO (SWAPPED DOWN - HIGHER PRIORITY) ---
            console.log(`STEP 9: Applying PTO for ${sheetName}...`);
            
            // 1. Load PTO Named Ranges
            const rngWho = context.workbook.names.getItemOrNullObject("Who").getRangeOrNullObject();
            const rngStart = context.workbook.names.getItemOrNullObject("StartDate").getRangeOrNullObject();
            const rngDays = context.workbook.names.getItemOrNullObject("NumberDays").getRangeOrNullObject();

            rngWho.load(["values", "isNullObject"]);
            rngStart.load(["values", "isNullObject"]);
            rngDays.load(["values", "isNullObject"]);
            await context.sync();

            if (!rngWho.isNullObject && rngWho.values) {
                const ptoNames = rngWho.values;
                const ptoStarts = rngStart.values;
                const ptoDays = rngDays.values;

                for (let i = 0; i < ptoNames.length; i++) {
                    const name = ptoNames[i][0]?.toString().toUpperCase().trim();
                    const start = ptoStarts[i][0];
                    const days = ptoDays[i][0];

                    // Only apply if the person belongs to THIS office
                    if (name && officeMap[name] === sheetName && typeof start === 'number') {
                        const end = start + days - 1;
                        const fPTO = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                        fPTO.custom.rule.formula = `=AND(K$6>=${start}, K$6<=${end})`;
                        fPTO.custom.format.fill.color = "#EAEAEA";
                        fPTO.stopIfTrue = false;

                        // Batch sync to avoid request overflow
                        if (i % 5 === 0) await context.sync();
                    }
                }
            }
            await context.sync();
            console.log(">> PTO Applied.");

            console.log("Formatting Engine: ALL RULES SUCCESSFUL.");

        } catch (error) {
            console.error("Formatting Logic Error at Step:", error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug Info:", error.debugInfo);
            }
        } finally {
            // 2. HIDE SPINNER (Always runs, even if error)
            if (window.GlobalLoader) window.GlobalLoader.hide();
        }
    }
};
