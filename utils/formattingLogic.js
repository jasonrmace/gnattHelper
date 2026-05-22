/* global Excel */

// ==============================================================================
// FORMATTING LOGIC ENGINE (Phase 5: Corrected Stack & Filters)
// Replaces VBA 'GenerateSmartRules'
// 
// VISUAL STACK (Logic Order):
// 1. Progress Bar (Grey Overlay) -> Added FIRST
// 2. Today Borders (Red Outline)
// 3. Team Colors (Specific Assignees)
// 4. Generic Blue (Default Task)
// 5. Parent Row (Grey Row)
// 6. Holidays (Grey Background)
// 7. PTO (Grey Background) -> Added LAST
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

            // 3. LOAD TEAM COLORS (Safe Mode)
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
                if (colors[i] && colors[i][0]) {
                    let hex = colors[i][0].format.fill.color;
                    
                    // FILTER: Case-Insensitive Check for White/Null
                    // If the user has no color, we MUST skip this rule so the Generic Blue rule applies.
                    let isInvalid = !hex || hex === "null";
                    if (typeof hex === 'string' && hex.toUpperCase() === "#FFFFFF") isInvalid = true;
                    
                    if (name && !isInvalid) { 
                        teamRules.push({ 
                            name: name.toString().toUpperCase().trim(), 
                            color: hex 
                        });
                    }
                }
            }
            console.log(`Formatting: Loaded ${teamRules.length} color rules.`);

            // =================================================================
            // APPLY RULES (Order: FIRST ADDED = TOP PRIORITY)
            // =================================================================

            // --- 1. PROGRESS BAR (Top Layer) ---
            // Greys out the COMPLETED portion of the bar.
            const fProg = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fProg.custom.rule.formula = '=AND(K$6>=$E8, K$6<=$F8, ISNUMBER($H8), ((K$6-$E8)/($F8-$E8+1)) < $H8)';
            fProg.custom.format.fill.color = "#D9D9D9"; // Grey 217
            fProg.stopIfTrue = false; // Allow borders/text to show through if needed

            // --- 2. TODAY BORDERS ---
            // Red borders on Left/Right. NO FILL.
            const fToday = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fToday.custom.rule.formula = '=K$6=TODAY()';
            fToday.custom.format.borders.getItem("Left").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Left").color = "#FF0000";
            fToday.custom.format.borders.getItem("Left").weight = Excel.BorderWeight.thick;
            fToday.custom.format.borders.getItem("Right").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Right").color = "#FF0000";
            fToday.custom.format.borders.getItem("Right").weight = Excel.BorderWeight.thick;
            fToday.stopIfTrue = false; // CRITICAL: Must be FALSE so the Blue/Team color below shows through!

            // --- 3. TEAM COLORS (Specific Overrides) ---
            // Only added for people who explicitly have a color (not white).
            for (let member of teamRules) {
                const safeName = member.name.replace(/"/g, '""');

                // Grid Bar
                const fBar = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fBar.custom.rule.formula = `=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8, UPPER(TRIM($C8))="${safeName}")`;
                fBar.custom.format.fill.color = member.color;
                fBar.stopIfTrue = true; // Stop here. Do not fall through to Blue.

                // Name Cell
                const fName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fName.custom.rule.formula = `=UPPER(TRIM($C8))="${safeName}"`;
                fName.custom.format.fill.color = member.color;
                fName.stopIfTrue = true;
            }

            // --- 4. GENERIC BLUE (Default Task) ---
            // Catches everyone else (White/Null/No-Match).
            const fBlue = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fBlue.custom.rule.formula = '=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8)'; 
            fBlue.custom.format.fill.color = "#0070C0"; // Blue 192
            fBlue.stopIfTrue = true; // Stop here. Do not show Holidays underneath.

            // --- 5. PARENT ROW BACKGROUND ---
            const fParent = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParent.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParent.custom.format.fill.color = "#D9D9D9"; 
            fParent.stopIfTrue = true; 
            
            const fParentName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParentName.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParentName.custom.format.fill.color = "#D9D9D9";
            fParentName.stopIfTrue = true;

            // --- 6. HOLIDAYS ---
            const fHol = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fHol.custom.rule.formula = '=COUNTIF(ListHolidays,K$6)>0';
            fHol.custom.format.fill.color = "#C8C8C8"; // Grey 200
            fHol.stopIfTrue = false;

            // --- 7. PTO (Bottom Priority) ---
            const fPTO = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fPTO.custom.rule.formula = '=SUMPRODUCT((Who=$C8)*(StartDate<=K$6)*((StartDate+NumberDays-1)>=K$6))>0';
            fPTO.custom.format.fill.color = "#EAEAEA"; // Grey 234
            fPTO.stopIfTrue = false;

            await context.sync();
            console.log("Formatting Engine: Rules Applied Successfully.");

        } catch (error) {
            console.error("Formatting Logic Error:", error);
        }
    }
};
