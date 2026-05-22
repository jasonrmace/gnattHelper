/* global Excel */

// ==============================================================================
// VISUAL LOGIC ENGINE (Phase 3 Migration)
// Replaces VBA 'UpdateGrid'
// Handles Holidays & PTO Data Validation Popups
// ==============================================================================

window.VisualLogic = {
    refreshGridAlerts: async (context) => {
        console.log("Visual Logic: Starting Grid Refresh...");
        try {
            const sheet = context.workbook.worksheets.getItem("GanttChart");
            
            // 1. LOAD CONFIG DATA (Holidays & PTO)
            // We use safe loading in case ranges are missing
            let holidayData = [];
            let ptoData = { names: [], starts: [], days: [] };

            // A. Load Holidays Table
            const holidayTable = sheet.tables.getItemOrNullObject("Holidays");
            await context.sync();
            
            if (!holidayTable.isNullObject) {
                const holidayBody = holidayTable.getDataBodyRange();
                holidayBody.load("values");
                await context.sync();
                // Assuming Col 1 = Date, Col 2 = Name
                holidayData = holidayBody.values.map(row => ({
                    date: row[0], // Excel Serial Date
                    name: row[1]
                }));
            }

            // B. Load PTO Named Ranges
            const rngWho = context.workbook.names.getItemOrNullObject("Who").getRangeOrNullObject();
            const rngStart = context.workbook.names.getItemOrNullObject("StartDate").getRangeOrNullObject();
            const rngDays = context.workbook.names.getItemOrNullObject("NumberDays").getRangeOrNullObject();
            
            rngWho.load(["values", "isNullObject"]);
            rngStart.load("values");
            rngDays.load("values");
            await context.sync();

            if (!rngWho.isNullObject && rngWho.values) {
                ptoData.names = rngWho.values.flat();
                ptoData.starts = rngStart.values.flat();
                ptoData.days = rngDays.values.flat();
            }

            // 2. SETUP GRID DIMENSIONS
            // Find Footer to determine height
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();

            const startRow = 7; // Row 8
            const endRow = footerRange.rowIndex - 1;
            const rowCount = endRow - startRow + 1;
            
            // Find Date Headers (Row 6, Col K onwards)
            // We limit to 365 columns (~1 year) to prevent timeouts
            const dateRowIndex = 5; // Row 6
            const startColIndex = 10; // Col K
            const colCount = 365; 

            const headerRange = sheet.getRangeByIndexes(dateRowIndex, startColIndex, 1, colCount);
            headerRange.load("values");
            await context.sync();
            
            const dates = headerRange.values[0];

            // 3. BATCH UPDATE COLUMNS
            // We iterate columns, determine the status, and apply validation to the WHOLE column strip
            
            for (let c = 0; c < dates.length; c++) {
                const serialDate = dates[c];
                if (typeof serialDate !== 'number') continue; // Skip empty/text headers

                let title = "";
                let message = "";
                let type = "none"; // 'holiday', 'pto', 'none'

                // Check Holiday
                const hol = holidayData.find(h => h.date === serialDate);
                if (hol) {
                    type = "holiday";
                    title = "Holiday";
                    message = hol.name;
                } 
                else {
                    // Check PTO
                    // Collect ALL people on PTO this day
                    const ptoNames = [];
                    for (let i = 0; i < ptoData.names.length; i++) {
                        const start = ptoData.starts[i];
                        const duration = ptoData.days[i];
                        const end = start + duration - 1;
                        
                        if (serialDate >= start && serialDate <= end) {
                            ptoNames.push(ptoData.names[i]);
                        }
                    }

                    if (ptoNames.length > 0) {
                        type = "pto";
                        title = "PTO Alert";
                        const suffix = ptoNames.length === 1 ? " is on PTO." : " are on PTO.";
                        message = ptoNames.join(", ") + suffix;
                    }
                }

                // Apply to Column Strip
                const colStrip = sheet.getRangeByIndexes(startRow, startColIndex + c, rowCount, 1);

                if (type !== "none") {
                    colStrip.dataValidation.rule = {
                        inputMessage: {
                            title: title,
                            message: message
                        }
                    };
                } else {
                    // Clear if standard workday (Optional: Checking first saves time, but clearing ensures accuracy)
                    colStrip.dataValidation.clear();
                }
                
                // Sync every 50 columns to keep Excel responsive
                if (c % 50 === 0) await context.sync();
            }

            await context.sync();
            console.log("Visual Logic: Grid Alerts Updated.");
            
        } catch (error) {
            console.error("VisualLogic Error:", error);
        }
    }
};
