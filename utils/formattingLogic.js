/* global Excel */

// ==============================================================================
// FORMATTING LOGIC ENGINE (Phase 4 Final - FIFO Order)
// Replaces VBA 'GenerateSmartRules'
// Strategy: 
// 1. EXECUTION ORDER = VISUAL PRIORITY ORDER (Top to Bottom).
// 2. Progress Bar & Today Borders go FIRST (Top).
// 3. PTO & Holidays go LAST (Bottom).
// ==============================================================================

window.FormattingLogic = {
    generateSmartRules: async (context) => {
        console.log("Formatting Engine: Starting...");
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
            const colCount = headerRange.columnCount;

            // Ranges
            const gridRange = sheet.getRangeByIndexes(startRow, 10, rowCount, colCount); // Col K
            const namesRange = sheet.getRangeByIndexes(startRow, 2, rowCount, 1); // Col C

            // 2. CLEAN SLATE
            gridRange.conditionalFormats.clearAll();
            namesRange.conditionalFormats.clearAll();

            // 3. LOAD TEAM COLORS
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
                const name = names[i][0];
                // Safe Access
                if (colors[i] && colors[i][0]) {
                    let hex = colors[i][0].format.fill.color;
                    
                    // FILTER: If White/Null, we SKIP adding a rule.
                    // This lets the row "fall through" to the Generic Blue rule (Layer 4).
                    const isInvalid = !hex || hex === "null" || hex === "#FFFFFF";
                    
                    if (name && !isInvalid) { 
                        teamRules.push({ 
                            name: name.toString().toUpperCase().trim(), 
                            color: hex 
                        });
                    }
                }
            }

            // =================================================================
            // APPLY RULES (Execution Order: TOP -> BOTTOM)
            // First Rule Added = Priority #1
            // =================================================================

            // --- PRIORITY 1: PROGRESS BAR (Overlay) ---
            // Must be top to grey out the colors below it
            const fProg = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fProg.custom.rule.formula = '=AND(K$6>=$E8, K$6<=$F8, ISNUMBER($H8), ((K$6-$E8)/($F8-$E8+1)) < $H8)';
            fProg.custom.format.fill.color = "#D9D9D9";
            fProg.stopIfTrue = false; // Allow seeing through to next layers if needed, but it's opaque grey.

            // --- PRIORITY 2: TODAY BORDERS ---
            // Sits on top of the Blue/Team bars.
            const fToday = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fToday.custom.rule.formula = '=K$6=TODAY()';
            fToday.custom.format.borders.getItem("Left").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Left").color = "#FF0000";
            fToday.custom.format.borders.getItem("Left").weight = Excel.BorderWeight.thick; // Make it visible
            fToday.custom.format.borders.getItem("Right").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Right").color = "#FF0000";
            fToday.custom.format.borders.getItem("Right").weight = Excel.BorderWeight.thick;
            fToday.stopIfTrue = false; // CRITICAL: Must be false so the Fill Color (Blue/Team) shows through!

            // --- PRIORITY 3: TEAM COLORS ---
            // Specific overrides for people with assigned colors
            for (let member of teamRules) {
                const safeName = member.name.replace(/"/g, '""');

                // Grid Bar
                const fBar = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fBar.custom.rule.formula = `=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8, UPPER(TRIM($C8))="${safeName}")`;
                fBar.custom.format.fill.color = member.color;
                fBar.stopIfTrue = true; // Stop here. Don't turn Blue.

                // Name Cell
                const fName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fName.custom.rule.formula = `=UPPER(TRIM($C8))="${safeName}"`;
                fName.custom.format.fill.color = member.color;
                fName.stopIfTrue = true;
            }

            // --- PRIORITY 4: GENERIC BLUE (Fallback) ---
            // Everyone else (including White/Null colors) gets this.
            const fBlue = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fBlue.custom.rule.formula = '=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8)'; 
            fBlue.custom.format.fill.color = "#0070C0"; 
            fBlue.stopIfTrue = true; // Stop here. Don't show Holidays/PTO underneath the bar.

            // --- PRIORITY 5: PARENT ROW ---
            const fParent = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParent.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParent.custom.format.fill.color = "#D9D9D9"; 
            fParent.stopIfTrue = true; 
            
            const fParentName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParentName.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParentName.custom.format.fill.color = "#D9D9D9";
            fParentName.stopIfTrue = true;

            // --- PRIORITY 6: HOLIDAYS ---
            const fHol = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fHol.custom.rule.formula = '=COUNTIF(ListHolidays,K$6)>0';
            fHol.custom.format.fill.color = "#C8C8C8"; 
            fHol.stopIfTrue = false;

            // --- PRIORITY 7: PTO (Bottom) ---
            const fPTO = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fPTO.custom.rule.formula = '=SUMPRODUCT((Who=$C8)*(StartDate<=K$6)*((StartDate+NumberDays-1)>=K$6))>0';
            fPTO.custom.format.fill.color = "#EAEAEA";
            fPTO.stopIfTrue = false;

            await context.sync();
            console.log("Formatting Engine: Rules Applied Successfully (FIFO Order).");

        } catch (error) {
            console.error("Formatting Logic Error:", error);
        }
    }
};
