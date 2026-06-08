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
            
            // SIGNAL START: Write to a global flag so co-authors see the spinner
            const statusRange = context.workbook.names.getItemOrNullObject("GlobalFormattingStatus").getRangeOrNullObject();
            statusRange.load("isNullObject");
            await context.sync();
            if (!statusRange.isNullObject) statusRange.values = [["IN_PROGRESS"]];

            // 1. DEFINE BOUNDARIES
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
            
            // Dynamically find the last column with a header in Row 7 (Index 6)
            const headerRow = sheet.getRange("7:7");
            const usedHeaderRange = headerRow.getUsedRangeOrNullObject();
            usedHeaderRange.load(["columnCount", "columnIndex", "isNullObject"]);
            
            // Fetch all Team/PTO metadata once to avoid redundant reads in the row loop
            const metadata = await FormattingLogic.fetchFormattingMetadata(context);
            await context.sync();
            
            let lastColIndex = 110; // Default to Column DF if empty
            if (!usedHeaderRange.isNullObject) {
                lastColIndex = usedHeaderRange.columnIndex + usedHeaderRange.columnCount - 1;
            }

            const startColIndex = 10; // Start at Column K
            let colCount = Math.max(lastColIndex - startColIndex + 1, 10);
            if (colCount > 500) colCount = 500; // Performance safety cap

            console.log(`Resetting Worksheet: ${sheetName} (${actualRowCount} rows, ${colCount} columns)...`);

            // 2. CLEAR ALL PREVIOUS CF RULES IN TARGET AREA
            const gridRange = sheet.getRangeByIndexes(startRow, 10, actualRowCount, colCount);
            const namesRange = sheet.getRangeByIndexes(startRow, 2, actualRowCount, 1);
            gridRange.conditionalFormats.clearAll();
            namesRange.conditionalFormats.clearAll();
            await context.sync();

            // 3. APPLY ROW BY ROW (As Requested for surgical precision)
            for (let i = 0; i < actualRowCount; i++) {
                const currentRow = startRow + i;
                if (window.GlobalLoader) window.GlobalLoader.show(`Formatting ${sheetName}: Row ${i + 1}/${actualRowCount}...`);
                await FormattingLogic.applyRulesToRange(context, sheetName, currentRow, 1, colCount, metadata);
            }

        } catch (error) {
            console.error("Formatting Logic Error:", error);
        } finally {
            // SIGNAL END: Clear the global flag
            await Excel.run(async (ctx) => {
                const statusRange = ctx.workbook.names.getItemOrNullObject("GlobalFormattingStatus").getRangeOrNullObject();
                statusRange.load("isNullObject");
                await ctx.sync();
                if (!statusRange.isNullObject) statusRange.values = [["IDLE"]];
                await ctx.sync();
            }).catch(() => {});

            if (window.GlobalLoader) window.GlobalLoader.hide();
        }
    },

    /**
     * Fetches all necessary data from 'Team' and PTO ranges to avoid 
     * redundant reads during row-by-row updates.
     */
    fetchFormattingMetadata: async (context) => {
        const teamSheet = context.workbook.worksheets.getItemOrNullObject("Team");
        const rngWho = context.workbook.names.getItemOrNullObject("Who").getRangeOrNullObject();
        const rngStart = context.workbook.names.getItemOrNullObject("StartDate").getRangeOrNullObject();
        const rngDays = context.workbook.names.getItemOrNullObject("NumberDays").getRangeOrNullObject();

        teamSheet.load("isNullObject");
        rngWho.load(["values", "isNullObject"]);
        rngStart.load(["values", "isNullObject"]);
        rngDays.load(["values", "isNullObject"]);
        
        let officeMap = {};
        let teamRules = [];
        let ptoData = [];

        const teamTable = teamSheet.tables.getItemOrNullObject("Team");
        teamTable.load("isNullObject");
        await context.sync();

        if (!teamTable.isNullObject) {
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
                const rawName = names[i][0]?.toString().toUpperCase().trim();
                if (rawName) {
                    officeMap[rawName] = offices[i][0];
                    let hex = colors[i][0]?.format.fill.color;
                    if (hex && hex !== "null" && hex.toUpperCase() !== "#FFFFFF") {
                        teamRules.push({ name: rawName, color: hex });
                    }
                }
            }
        }

        if (!rngWho.isNullObject && rngWho.values) {
            ptoData = rngWho.values.map((v, i) => ({
                name: v[0]?.toString().toUpperCase().trim(),
                start: rngStart.values[i][0],
                days: rngDays.values[i][0]
            }));
        }

        return { teamRules, officeMap, ptoData };
    },

    /**
     * Applies logic to a SPECIFIC range without clearing existing rules elsewhere.
     * @param {Object} metadata Optional cached metadata to speed up row-level updates.
     */
    applyRulesToRange: async (context, sheetName, startRow, rowCount, colCount, metadata = null) => {
        try {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            if (!metadata) metadata = await FormattingLogic.fetchFormattingMetadata(context);

            const gridRange = sheet.getRangeByIndexes(startRow, 10, rowCount, colCount);
            const namesRange = sheet.getRangeByIndexes(startRow, 2, rowCount, 1);
            const r = startRow + 1;

            if (rowCount <= 0) return;

            // Surgical cleanup for the targeted row(s)
            gridRange.unmerge();
            namesRange.unmerge();
            gridRange.format.fill.clear();
            namesRange.format.fill.clear();
            
            gridRange.conditionalFormats.clearAll();
            namesRange.conditionalFormats.clearAll();

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

            // --- BLOCK A: STANDARD PROGRESS BAR (PARTIAL) ---
            const fProg = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fProg.custom.rule.formula = `=AND(K$6>=$E${r}, K$6<=$F${r}, ISNUMBER($H${r}), ((K$6-$E${r})/($F${r}-$E${r}+1)) < $H${r})`;
            fProg.custom.format.fill.color = "#5a5a5a"; // Dark Grey (User Choice)
            fProg.stopIfTrue = false;

            // --- BLOCK B: TODAY BORDERS ---
            const fToday = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fToday.custom.rule.formula = '=K$6=TODAY()';
            
            fToday.custom.format.borders.getItem("EdgeLeft").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("EdgeLeft").color = "#FF0000";
            fToday.custom.format.borders.getItem("EdgeLeft").weight = Excel.BorderWeight.thick;
            
            fToday.custom.format.borders.getItem("EdgeRight").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("EdgeRight").color = "#FF0000";
            fToday.custom.format.borders.getItem("EdgeRight").weight = Excel.BorderWeight.thick;
            fToday.stopIfTrue = false;

            // --- BLOCK C: TEAM COLORS (PROJECT PROTECTED) ---
            if (metadata.teamRules.length > 0) {
                for (const member of metadata.teamRules) {
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
                }
            }

            // --- BLOCK D: GENERIC BLUE ---
            const fBlue = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fBlue.custom.rule.formula = `=AND(ISNUMBER($E${r}), K$6>=$E${r}, K$6<=$F${r})`; 
            fBlue.custom.format.fill.color = "#0070C0"; 
            fBlue.stopIfTrue = true; 

            // --- BLOCK E: PARENT ROW ---
            const fParent = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
fParent.custom.rule.formula = `=AND($A${r}<>"", ISERROR(SEARCH(".", $A${r})))`;
            fParent.custom.format.fill.color = "#D9D9D9"; 
            fParent.stopIfTrue = true; 
            
            const fParentName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParentName.custom.rule.formula = `=AND($A${r}<>"", ISERROR(SEARCH(".", $A${r})))`;
            fParentName.custom.format.fill.color = "#D9D9D9";
            fParentName.stopIfTrue = true;

            // --- BLOCK G: HOLIDAYS (SWAPPED UP) ---
            // Using IFERROR/ISERROR safety in case ListHolidays named range doesn't exist yet
            const fHol = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fHol.custom.rule.formula = '=IFERROR(COUNTIF(ListHolidays,K$6)>0, FALSE)';
            fHol.custom.format.fill.color = "#C8C8C8"; 
            fHol.stopIfTrue = false;

            // --- BLOCK F: PTO (SWAPPED DOWN - HIGHER PRIORITY) ---
            for (const pto of metadata.ptoData) {
                if (pto.name && metadata.officeMap[pto.name] === sheetName && typeof pto.start === 'number') {
                    const end = pto.start + pto.days - 1;
                    const fPTO = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                    fPTO.custom.rule.formula = `=AND(K$6>=${pto.start}, K$6<=${end})`;
                    fPTO.custom.format.fill.color = "#EAEAEA";
                    fPTO.stopIfTrue = false;
                }
            }
            
            await context.sync();

        } catch (error) {
            console.error("Formatting Logic Error at Step:", error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug Info:", error.debugInfo);
            }
        }
    }
};
