/* global Excel */

// ==============================================================================
// FORMATTING LOGIC ENGINE (Phase 4 Final - VBA Match)
// Replaces VBA 'GenerateSmartRules'
// Strategy: 
// 1. Adds rules from BOTTOM to TOP (Last Added = Priority 1).
// 2. Filters out White/Null colors so they default to Blue (matching VBA).
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
                    
                    // LOGIC MATCHING VBA: 
                    // If Color is White (#FFFFFF) or Null, VBA converts it to Blue.
                    // In JS, we achieve this by SKIPPING the specific rule and letting
                    // the "Generic Blue" rule (Layer 4) handle it.
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
            // APPLY RULES (Execution Order: BOTTOM -> TOP)
            // The LAST rule added becomes Priority #1 (Top Visual Layer).
            // This recreates the "Total Priority Sort" from VBA.
            // =================================================================

            // --- LAYER 1: PTO (Absolute Bottom) ---
            const fPTO = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fPTO.custom.rule.formula = '=SUMPRODUCT((Who=$C8)*(StartDate<=K$6)*((StartDate+NumberDays-1)>=K$6))>0';
            fPTO.custom.format.fill.color = "#EAEAEA"; // Grey 234
            fPTO.stopIfTrue = false;

            // --- LAYER 2: HOLIDAYS ---
            const fHol = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fHol.custom.rule.formula = '=COUNTIF(ListHolidays,K$6)>0';
            fHol.custom.format.fill.color = "#C8C8C8"; // Grey 200
            fHol.stopIfTrue = false;

            // --- LAYER 3: PARENT ROW BACKGROUND ---
            const fParent = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParent.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParent.custom.format.fill.color = "#D9D9D9"; 
            fParent.stopIfTrue = false; 
            
            const fParentName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fParentName.custom.rule.formula = '=AND($A8<>"", ISERROR(SEARCH(".", $A8)))';
            fParentName.custom.format.fill.color = "#D9D9D9";
            fParentName.stopIfTrue = false;

            // --- LAYER 4: GENERIC BLUE (Default for Tasks) ---
            // This acts as the "Background" for tasks. 
            // If a Team Color exists, it sits ON TOP of this (Layer 5).
            const fBlue = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fBlue.custom.rule.formula = '=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8)'; 
            fBlue.custom.format.fill.color = "#0070C0"; // Blue 192
            fBlue.stopIfTrue = false; 

            // --- LAYER 5: TEAM COLORS (Specific Overrides) ---
            // Added AFTER Blue, so they sit ON TOP of Blue.
            // We only add rules for people who actually have a color (Not White).
            for (let member of teamRules) {
                const safeName = member.name.replace(/"/g, '""');

                // Grid Bar
                const fBar = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fBar.custom.rule.formula = `=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8, UPPER(TRIM($C8))="${safeName}")`;
                fBar.custom.format.fill.color = member.color;
                fBar.stopIfTrue = false;

                // Name Cell
                const fName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fName.custom.rule.formula = `=UPPER(TRIM($C8))="${safeName}"`;
                fName.custom.format.fill.color = member.color;
                fName.stopIfTrue = false;
            }

            // --- LAYER 6: TODAY BORDERS ---
            // Added AFTER Colors, so the Red Borders draw ON TOP of the Blue/Team bars.
            const fToday = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fToday.custom.rule.formula = '=K$6=TODAY()';
            fToday.custom.format.borders.getItem("Left").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Left").color = "#FF0000";
            fToday.custom.format.borders.getItem("Right").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Right").color = "#FF0000";
            fToday.stopIfTrue = false; 

            // --- LAYER 7: PROGRESS BAR (Top Overlay) ---
            // Added LAST, so it overlays everything else (Visual Priority #1)
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
