/* global Excel */

// ==============================================================================
// VISUAL LOGIC ENGINE (Phase 10: English Grammar Fix)
// Purpose: Applies Data Validation (Popups) AND In-Cell Counters.
// 
// CHANGE LOG:
// 1. Formatting: Now uses "&" for the last name (e.g., "A, B & C are on PTO").
// ==============================================================================

window.VisualLogic = {
    refreshGridAlerts: async (context) => {
        console.log("Visual Logic: Starting Grid Refresh (Phase 10 - Grammar)...");
        
        // 1. SHOW SPINNER
        if (window.GlobalLoader) window.GlobalLoader.show("Syncing Visuals...");

        try {
            const sheet = context.workbook.worksheets.getItem("GanttChart");

            // 1. LOAD CONFIG DATA (Holidays & PTO)
            let holidayData = [];
            let ptoData = { names: [], starts: [], days: [] };

            // A. Load Holidays Table
            const holidayTable = sheet.tables.getItemOrNullObject("Holidays");
            await context.sync();

            if (!holidayTable.isNullObject) {
                holidayTable.load("rows/count");
                await context.sync();

                if (holidayTable.rows.count > 0) {
                    const holidayBody = holidayTable.getDataBodyRange();
                    holidayBody.load("values");
                    await context.sync();
                    holidayData = holidayBody.values.map(row => ({ 
                        date: row[0], 
                        name: row[1] 
                    }));
                }
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
                ptoData.names = rngWho.values.map(v => v[0]);
                ptoData.starts = rngStart.values.map(v => v[0]);
                ptoData.days = rngDays.values.map(v => v[0]);
            }

            // 2. SETUP GRID DIMENSIONS
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();

            const startRow = 7; // Row 8
            const endRow = footerRange.rowIndex - 1;
            const rowCount = endRow - startRow + 1;

            // Find Date Headers
            const dateRowIndex = 5; // Row 6
            const startColIndex = 10; // Col K
            const colCount = 365; 

            const headerRange = sheet.getRangeByIndexes(dateRowIndex, startColIndex, 1, colCount);
            headerRange.load("values");
            await context.sync();

            const dates = headerRange.values[0];

            // 3. BATCH UPDATE COLUMNS
            console.log(`Visual Logic: Scanning ${colCount} days...`);

            for (let c = 0; c < dates.length; c++) {
                const serialDate = dates[c];
                
                if (typeof serialDate === 'number') {
                    let title = "";
                    let message = "";
                    let type = "none"; 
                    let ptoCount = 0;

                    // Check Holiday
                    const hol = holidayData.find(h => h.date === serialDate);
                    if (hol) {
                        type = "holiday";
                        title = "Holiday";
                        message = hol.name;
                    } else {
                        // Check PTO
                        const ptoNames = [];
                        for (let i = 0; i < ptoData.names.length; i++) {
                            const start = ptoData.starts[i];
                            const duration = ptoData.days[i];
                            
                            if (typeof start === 'number' && typeof duration === 'number') {
                                const end = start + duration - 1;
                                if (serialDate >= start && serialDate <= end) {
                                    ptoNames.push(ptoData.names[i]);
                                }
                            }
                        }

                        if (ptoNames.length > 0) {
                            type = "pto";
                            title = "PTO Alert";
                            ptoCount = ptoNames.length;
                            const suffix = ptoCount === 1 ? " is on PTO." : " are on PTO.";
                            
                            // --- GRAMMAR FIX ---
                            if (ptoCount === 1) {
                                message = ptoNames[0] + suffix;
                            } else if (ptoCount === 2) {
                                message = `${ptoNames[0]} & ${ptoNames[1]}` + suffix;
                            } else {
                                const last = ptoNames.pop(); // Remove last name
                                message = `${ptoNames.join(", ")} & ${last}` + suffix;
                            }
                        }
                    }

                    // --- APPLY TO COLUMN STRIP ---
                    const colStrip = sheet.getRangeByIndexes(startRow, startColIndex + c, rowCount, 1);

                    // A. Apply Data Validation (Popup)
                    colStrip.dataValidation.clear();
                    if (type !== "none") {
                        if (message.length > 255) message = message.substring(0, 250) + "...";
                        colStrip.dataValidation.prompt = {
                            showPrompt: true,
                            title: title,
                            message: message
                        };
                    }

                    // B. Apply In-Cell Counter (Superscript)
                    let cellValue = "";
                    if (ptoCount >= 2) {
                        cellValue = ptoCount;
                    }

                    if (ptoCount >= 2) {
                        // 2+ People: Write the Number
                        const columnValues = new Array(rowCount).fill([ptoCount]);
                        colStrip.values = columnValues;
                        
                        // Format: Center, Dark Grey, Superscript
                        colStrip.format.font.superscript = true;
                        colStrip.format.horizontalAlignment = "Center";
                        colStrip.format.font.color = "#444444"; 
                    } else {
                        // 0 or 1 Person: Clear the cell contents
                        colStrip.clear(Excel.ClearApplyTo.contents);
                    }
                }

                // Sync every 50 columns
                if (c % 50 === 0) await context.sync();
            }

            await context.sync();
            console.log("Visual Logic: Grid Updated.");

        } catch (error) {
            console.error("VisualLogic Error:", error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug Info:", error.debugInfo);
            }
        } finally {
            // 2. HIDE SPINNER
            if (window.GlobalLoader) window.GlobalLoader.hide();
        }
    }
};
