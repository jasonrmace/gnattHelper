/* global Excel */

// ==============================================================================
// FORMATTING LOGIC ENGINE (Phase 4 Fix - Attempt 2)
// Replaces VBA 'GenerateSmartRules'
// Fixes: 
// 1. Corrects 2D Array access for Team Colors (colors[i][0])
// 2. Explicitly sets stopIfTrue = false for correct layering
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

            // The Main Grid & Names Column
            const gridRange = sheet.getRangeByIndexes(startRow, 10, rowCount, colCount); // Col K
            const namesRange = sheet.getRangeByIndexes(startRow, 2, rowCount, 1); // Col C

            // 2. CLEAN SLATE
            gridRange.conditionalFormats.clearAll();
            namesRange.conditionalFormats.clearAll();

            // 3. LOAD TEAM COLORS (FIXED)
            let teamRules = [];
            const teamTable = teamSheet.tables.getItem("Team");
            const nameCol = teamTable.columns.getItem("First Name").getDataBodyRange();
            const colorCol = teamTable.columns.getItem("Color").getDataBodyRange();
            
            nameCol.load("values");
            // USE getCellProperties to get individual colors
            const colorProps = colorCol.getCellProperties({ format: { fill: { color: true } } });
            
            await context.sync();

            const names = nameCol.values;
            const colors = colorProps.value; // 2D Array: [Row][Col]

            // Map names to colors
            for (let i = 0; i < names.length; i++) {
                const name = names[i][0];
                // FIX: Access [0] because 'colors[i]' is a row array
                const hex = colors[i][0].format.fill.color; 
                
                if (name && hex && hex !== "#FFFFFF" && hex !== "null") { 
                    teamRules.push({ 
                        name: name.toString().toUpperCase().trim(), 
                        color: hex 
                    });
                }
            }

            // =================================================================
            // APPLY RULES (Stack Order: Last Added = Top Priority)
            // We assume 'add' pushes to top of stack.
            // =================================================================

            // --- LAYER 1: PARENT ROW (Bottom) ---
            const fParent = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParent.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParent.custom.format.fill.color = "#D9D9D9"; 
            fParent.stopIfTrue = false; // Critical for layering
            
            const fParentName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParentName.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParentName.custom.format.fill.color = "#D9D9D9";
            fParentName.stopIfTrue = false;

            // --- LAYER 2: HOLIDAYS ---
            const fHol = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fHol.custom.rule.formula = '=COUNTIF(ListHolidays,K$6)>0';
            fHol.custom.format.fill.color = "#C8C8C8";
            fHol.stopIfTrue = false;

            // --- LAYER 3: PTO ---
            const fPTO = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fPTO.custom.rule.formula = '=SUMPRODUCT((Who=$C8)*(StartDate<=K$6)*((StartDate+NumberDays-1)>=K$6))>0';
            fPTO.custom.format.fill.color = "#EAEAEA";
            fPTO.stopIfTrue = false;

            // --- LAYER 4: GENERIC BLUE (Fallback) ---
            const fBlue = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fBlue.custom.rule.formula = '=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8)'; 
            fBlue.custom.format.fill.color = "#0070C0";
            fBlue.stopIfTrue = false;

            // --- LAYER 5: TEAM COLORS (Dynamic - Sits ON TOP of Blue) ---
            for (let member of teamRules) {
                // Grid Bar
                const fBar = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                // Escaped quote for name safety
                fBar.custom.rule.formula = `=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8, UPPER(TRIM($C8))="${member.name.replace(/"/g, '""')}")`;
                fBar.custom.format.fill.color = member.color;
                fBar.stopIfTrue = false;

                // Name Cell
                const fName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fName.custom.rule.formula = `=UPPER(TRIM($C8))="${member.name.replace(/"/g, '""')}"`;
                fName.custom.format.fill.color = member.color;
                fName.stopIfTrue = false;
            }

            // --- LAYER 6: TODAY BORDER ---
            const fToday = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fToday.custom.rule.formula = '=K$6=TODAY()';
            fToday.custom.format.borders.getItem("Left").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Left").color = "#FF0000";
            fToday.custom.format.borders.getItem("Right").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Right").color = "#FF0000";
            fToday.stopIfTrue = false;

            // --- LAYER 7: PROGRESS BAR (Top Overlay) ---
            // Grey out completed portion
            const fProg = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fProg.custom.rule.formula = '=AND(K$6>=$E8, K$6<=$F8, ISNUMBER($H8), ((K$6-$E8)/($F8-$E8+1)) < $H8)';
            fProg.custom.format.fill.color = "#D9D9D9";
            fProg.stopIfTrue = false;

            await context.sync();
            console.log("Formatting Engine: Rules Applied Successfully.");

        } catch (error) {
            console.error("Formatting Logic Error:", error);
        }
    }
};
