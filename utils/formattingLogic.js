/* global Excel */

// ==============================================================================
// FORMATTING LOGIC ENGINE (Phase 4 Migration)
// Replaces VBA 'GenerateSmartRules'
// Handles Conditional Formatting (Bars, Colors, Holidays, Today Line)
// ==============================================================================

window.FormattingLogic = {
    generateSmartRules: async (context) => {
        console.log("Formatting Engine: Starting...");
        try {
            const sheet = context.workbook.worksheets.getItem("GanttChart");
            const teamSheet = context.workbook.worksheets.getItem("Team");

            // 1. DEFINE RANGES
            // Find Footer to get height
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();
            
            const startRow = 7; // Row 8
            const endRow = footerRange.rowIndex - 1;
            const rowCount = endRow - startRow + 1;
            
            // Find Date Extent (Col K to End)
            const headerRange = sheet.getRange("K6").getExtendedRange(Excel.KeyboardDirection.right);
            headerRange.load("columnCount");
            await context.sync();
            const colCount = headerRange.columnCount;

            // The Main Grid (Dates)
            const gridRange = sheet.getRangeByIndexes(startRow, 10, rowCount, colCount); // Col K (10)
            // The Names Column (Lead)
            const namesRange = sheet.getRangeByIndexes(startRow, 2, rowCount, 1); // Col C (2)

            // 2. CLEAN SLATE
            // We clear existing rules to prevent duplicates/conflicts
            gridRange.conditionalFormats.clearAll();
            namesRange.conditionalFormats.clearAll();

            // 3. LOAD TEAM COLORS
            let teamRules = [];
            const teamTable = teamSheet.tables.getItem("Team");
            const nameCol = teamTable.columns.getItem("First Name").getDataBodyRange();
            const colorCol = teamTable.columns.getItem("Color").getDataBodyRange();
            
            nameCol.load("values");
            colorCol.load("format/fill/color");
            await context.sync();

            // Map names to colors
            for (let i = 0; i < nameCol.values.length; i++) {
                const name = nameCol.values[i][0];
                const color = colorCol.format.fill.color;
                if (name && color) {
                    teamRules.push({ name: name.toString().toUpperCase().trim(), color: color });
                }
            }

            // =================================================================
            // APPLY RULES (Order Matters: First Applied = Bottom Layer)
            // In Office.js 'add' puts rule at the TOP. 
            // So we add HIGHEST PRIORITY LAST.
            // =================================================================

            // --- LAYER 1: PARENT ROW BACKGROUND (Bottom) ---
            // Formula: Row is integer (Parent)
            const fParent = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParent.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParent.custom.format.fill.color = "#D9D9D9"; // Grey 217
            
            const fParentName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParentName.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParentName.custom.format.fill.color = "#D9D9D9";

            // --- LAYER 2: HOLIDAYS ---
            const fHol = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fHol.custom.rule.formula = '=COUNTIF(ListHolidays,K$6)>0';
            fHol.custom.format.fill.color = "#C8C8C8"; // Grey 200

            // --- LAYER 3: PTO ---
            const fPTO = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            // Note: JS strings need escaped quotes
            fPTO.custom.rule.formula = '=SUMPRODUCT((Who=$C8)*(StartDate<=K$6)*((StartDate+NumberDays-1)>=K$6))>0';
            fPTO.custom.format.fill.color = "#EAEAEA"; // Grey 234

            // --- LAYER 4: GENERIC BLUE (Fallback) ---
            // Applied if name not in list
            const fBlue = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            // Use a simplified check for robustness or replicate exact VBA
            fBlue.custom.rule.formula = '=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8)'; 
            fBlue.custom.format.fill.color = "#0070C0"; // Standard Blue

            // --- LAYER 5: TEAM COLORS (Dynamic) ---
            // We loop and add them. Since we want them ON TOP of generic blue, we add them AFTER.
            for (let member of teamRules) {
                // Grid Bar
                const fBar = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fBar.custom.rule.formula = `=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8, UPPER(TRIM($C8))="${member.name}")`;
                fBar.custom.format.fill.color = member.color;

                // Name Cell
                const fName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fName.custom.rule.formula = `=UPPER(TRIM($C8))="${member.name}"`;
                fName.custom.format.fill.color = member.color;
            }

            // --- LAYER 6: TODAY BORDER (High Priority) ---
            const fToday = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fToday.custom.rule.formula = '=K$6=TODAY()';
            fToday.custom.format.borders.getItem("Left").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Left").color = "#FF0000"; // Red
            fToday.custom.format.borders.getItem("Right").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Right").color = "#FF0000";

            // --- LAYER 7: PROGRESS BAR (Highest Priority - Overlay) ---
            // It greys out the 'completed' portion of the bar
            // We assume Col H is % Done. Using Formula R1C1 or A1? A1 relative is tricky here.
            // Let's find the % column dynamically or assume H (Col 8)
            const fProg = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fProg.custom.rule.formula = '=AND(K$6>=$E8, K$6<=$F8, ISNUMBER($H8), ((K$6-$E8)/($F8-$E8+1)) < $H8)';
            fProg.custom.format.fill.color = "#D9D9D9"; // Grey Overlay

            await context.sync();
            console.log("Formatting Engine: Rules Applied.");

        } catch (error) {
            console.error("Formatting Logic Error:", error);
        }
    }
};
