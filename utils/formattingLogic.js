/* global Excel */

// ==============================================================================
// FORMATTING LOGIC ENGINE (Phase 6: Sync Fix & Range Safety)
// Replaces VBA 'GenerateSmartRules'
// Fixes: 
// 1. "Silent Failure" due to excessive column range (Caps at Header Width).
// 2. Race Condition: Syncs AFTER clear, BEFORE adding.
// 3. Explicitly targets K8:End based on headers.
// ==============================================================================

window.FormattingLogic = {
    generateSmartRules: async (context) => {
        console.log("Formatting Engine: Starting...");
        try {
            const sheet = context.workbook.worksheets.getItem("GanttChart");
            const teamSheet = context.workbook.worksheets.getItem("Team");

            // 1. DEFINE VERTICAL RANGE (ROWS)
            // From Row 8 (Index 7) down to the footer
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();
            
            const startRow = 7; // Row 8
            const endRow = footerRange.rowIndex - 1;
            const rowCount = endRow - startRow + 1;
            
            // 2. DEFINE HORIZONTAL RANGE (COLS)
            // We look at Row 6 (Headers) starting at K6 to find the width.
            // This matches VBA: ws.Cells(6, ws.Columns.Count).End(xlToLeft)
            const headerRange = sheet.getRange("K6").getExtendedRange(Excel.KeyboardDirection.right);
            headerRange.load("columnCount");
            await context.sync();
            
            let colCount = headerRange.columnCount;
            
            // SAFETY CAP: Prevent "Silent Failure" if range hits XFD (16,000 cols)
            // We cap at 2000 columns (approx 5.5 years of days), which is plenty but safe.
            if (colCount > 2000) {
                console.warn(`Col count ${colCount} is unsafe. Capping at 2000.`);
                colCount = 2000;
            }
            if (colCount < 1) colCount = 1;

            console.log(`Target Range: Rows=${rowCount}, Cols=${colCount} (Start: K8)`);

            // Define the Ranges
            const gridRange = sheet.getRangeByIndexes(startRow, 10, rowCount, colCount); // Col K (Index 10)
            const namesRange = sheet.getRangeByIndexes(startRow, 2, rowCount, 1); // Col C (Index 2)

            // 3. CLEAN SLATE & SYNC (CRITICAL FIX)
            // We MUST sync here to ensure the "Clear" finishes before we start adding.
            console.log("Clearing old rules...");
            gridRange.conditionalFormats.clearAll();
            namesRange.conditionalFormats.clearAll();
            await context.sync(); // <--- THIS PREVENTS THE SILENT FAILURE
            console.log("Rules cleared. Fetching colors...");

            // 4. LOAD TEAM COLORS
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
                    
                    // Case-Insensitive Check for White
                    if (typeof hex === 'string' && hex.toUpperCase() === "#FFFFFF") isInvalid = true;
                    if (!rawName) isInvalid = true;

                    if (!isInvalid) { 
                        teamRules.push({ 
                            name: rawName.toString().toUpperCase().trim(), 
                            color: hex 
                        });
                    }
                }
            }
            console.log(`Loaded ${teamRules.length} Team Rules.`);

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

            // --- 3. TEAM COLORS ---
            for (let member of teamRules) {
                const safeName = member.name.replace(/"/g, '""');
                
                // Name Cell Rule
                const fName = namesRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fName.custom.rule.formula = `=UPPER(TRIM($C8))="${safeName}"`;
                fName.custom.format.fill.color = member.color;
                fName.stopIfTrue = true;

                // Grid Bar Rule
                const fBar = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                fBar.custom.rule.formula = `=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8, UPPER(TRIM($C8))="${safeName}")`;
                fBar.custom.format.fill.color = member.color;
                fBar.stopIfTrue = true; 
            }

            // --- 4. GENERIC BLUE ---
            const fBlue = gridRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
            fBlue.custom.rule.formula = '=AND(ISNUMBER($E8), K$6>=$E8, K$6<=$F8)'; 
            fBlue.custom.format.fill.color = "#0070C0"; 
            fBlue.stopIfTrue = true; 

            // --- 5. PARENT ROW ---
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

            // --- 7. PTO ---
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
