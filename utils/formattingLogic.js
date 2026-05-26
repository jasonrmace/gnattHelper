/* global Excel */

// ==============================================================================
// FORMATTING LOGIC ENGINE (Phase 5: Debug & Verify)
// Replaces VBA 'GenerateSmartRules'
// Includes: Deep Logging to troubleshoot Missing Team Colors
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
            console.log("Clearing old rules...");
            gridRange.conditionalFormats.clearAll();
            namesRange.conditionalFormats.clearAll();

            // 3. LOAD TEAM COLORS (DEBUG MODE)
            let teamRules = [];
            console.log("Fetching Team Data...");
            
            const teamTable = teamSheet.tables.getItem("Team");
            const nameCol = teamTable.columns.getItem("First Name").getDataBodyRange();
            const colorCol = teamTable.columns.getItem("Color").getDataBodyRange();
            
            nameCol.load("values");
            const colorProps = colorCol.getCellProperties({ format: { fill: { color: true } } });
            await context.sync();

            const names = nameCol.values;
            const colors = colorProps.value;

            console.log(`found ${names.length} rows in Team table.`);

            for (let i = 0; i < names.length; i++) {
                const rawName = names[i][0];
                
                if (colors[i] && colors[i][0]) {
                    let hex = colors[i][0].format.fill.color;
                    
                    // --- DEBUG LOG FOR EACH ROW ---
                    // This will tell us EXACTLY what Excel is giving us
                    
                    let status = "KEEP";
                    let isInvalid = !hex || hex === "null";
                    
                    // Case-Insensitive Check for White
                    if (typeof hex === 'string' && hex.toUpperCase() === "#FFFFFF") isInvalid = true;
                    if (!rawName) isInvalid = true;

                    if (isInvalid) status = "SKIP (White/Empty)";

                    console.log(`Row ${i + 1}: Name="${rawName}" | Color="${hex}" | Action=${status}`);

                    if (status === "KEEP") { 
                        teamRules.push({ 
                            name: rawName.toString().toUpperCase().trim(), 
                            color: hex 
                        });
                    }
                }
            }
            console.log(`Formatting: Final list has ${teamRules.length} valid Color Rules.`);

            // =================================================================
            // APPLY RULES (Order: FIRST ADDED = TOP PRIORITY)
            // =================================================================

            // --- 1. PROGRESS BAR (Top Layer) ---
            const fProg = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fProg.custom.rule.formula = '=AND(K$6>=$E8, K$6<=$F8, ISNUMBER($H8), ((K$6-$E8)/($F8-$E8+1)) < $H8)';
            fProg.custom.format.fill.color = "#D9D9D9"; 
            fProg.stopIfTrue = false; 

            // --- 2. TODAY BORDERS ---
            const fToday = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fToday.custom.rule.formula = '=K$6=TODAY()';
            fToday.custom.format.borders.getItem("Left").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Left").color = "#FF0000";
            fToday.custom.format.borders.getItem("Left").weight = Excel.BorderWeight.thick;
            fToday.custom.format.borders.getItem("Right").style = Excel.BorderLineStyle.continuous;
            fToday.custom.format.borders.getItem("Right").color = "#FF0000";
            fToday.custom.format.borders.getItem("Right").weight = Excel.BorderWeight.thick;
            fToday.stopIfTrue = false;

            // --- 3. TEAM COLORS (Specific Overrides) ---
            if (teamRules.length === 0) {
                console.warn("WARNING: No Team Rules were generated! Everyone will be Blue.");
            }

            for (let member of teamRules) {
                const safeName = member.name.replace(/"/g, '""');
                console.log(`Applying Rule for: ${safeName} (${member.color})`);

                // Grid Bar
                const fBar = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fBar.custom.rule.formula = `=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8, UPPER(TRIM($C8))="${safeName}")`;
                fBar.custom.format.fill.color = member.color;
                fBar.stopIfTrue = true; 

                // Name Cell
                const fName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fName.custom.rule.formula = `=UPPER(TRIM($C8))="${safeName}"`;
                fName.custom.format.fill.color = member.color;
                fName.stopIfTrue = true;
            }

            // --- 4. GENERIC BLUE (Default Task) ---
            const fBlue = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fBlue.custom.rule.formula = '=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8)'; 
            fBlue.custom.format.fill.color = "#0070C0"; 
            fBlue.stopIfTrue = true; 

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
            fHol.custom.format.fill.color = "#C8C8C8"; 
            fHol.stopIfTrue = false;

            // --- 7. PTO (Bottom Priority) ---
            const fPTO = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fPTO.custom.rule.formula = '=SUMPRODUCT((Who=$C8)*(StartDate<=K$6)*((StartDate+NumberDays-1)>=K$6))>0';
            fPTO.custom.format.fill.color = "#EAEAEA"; 
            fPTO.stopIfTrue = false;

            await context.sync();
            console.log("Formatting Engine: Rules Applied Successfully.");

        } catch (error) {
            console.error("Formatting Logic Error:", error);
        }
    }
};
