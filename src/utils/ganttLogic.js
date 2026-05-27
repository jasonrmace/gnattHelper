/* global Excel */

// ==============================================================================
// GANTT LOGIC ENGINE (Phase 2 Migration)
// Replaces VBA 'UpdateParentAverages'
// ==============================================================================

export const GanttLogic = {
    updateProjectAverages: async (context) => {
        try {
            const sheet = context.workbook.worksheets.getItem("GanttChart");
            
            // 1. Find the Footer (Data Boundary)
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();
            
            const startRow = 7; // Row 8 (Index 7)
            const endRow = footerRange.rowIndex - 1;
            const rowCount = endRow - startRow + 1;
            
            if (rowCount <= 0) return;

            // 2. Batch Read: Columns A (IDs) and H (Formulas)
            // We read 'formulas' from H so we don't overwrite existing static values/formulas in child rows
            const rangeA = sheet.getRangeByIndexes(startRow, 0, rowCount, 1); // Col A
            const rangeH = sheet.getRangeByIndexes(startRow, 7, rowCount, 1); // Col H
            
            rangeA.load("text");
            rangeH.load("formulas");
            
            await context.sync();
            
            const valuesA = rangeA.text;
            const formulasH = rangeH.formulas;
            
            // 3. Calculate New Formulas in Memory
            // We modify 'formulasH' array in place, then write it back
            for (let i = 0; i < rowCount; i++) {
                const idStr = valuesA[i][0];
                const idNum = parseFloat(idStr);
                const currentRowIndex = startRow + i;

                // Check if this row is a PARENT (Integer, not empty)
                if (idStr && !isNaN(idNum) && Number.isInteger(idNum)) {
                    
                    let childStartRow = -1;
                    let childEndRow = -1;

                    // Look ahead for children
                    for (let j = i + 1; j < rowCount; j++) {
                        const nextIdStr = valuesA[j][0];
                        const nextIdNum = parseFloat(nextIdStr);

                        // Stop if we hit the NEXT Parent (Integer)
                        if (nextIdStr && !isNaN(nextIdNum) && Number.isInteger(nextIdNum)) {
                            break;
                        }
                        
                        // It is a child
                        if (childStartRow === -1) childStartRow = startRow + j;
                        childEndRow = startRow + j;
                    }

                    // If children found, update the formula for this Parent
                    if (childStartRow !== -1 && childEndRow !== -1) {
                        // Convert to Excel 1-based references (Index 7 = Row 8)
                        const refStart = childStartRow + 1;
                        const refEnd = childEndRow + 1;
                        formulasH[i] = [`=AVERAGE(H${refStart}:H${refEnd})`];
                    }
                }
                // If it's a Child, we leave formulasH[i] exactly as it was (preserving manual % or existing formulas)
            }

            // 4. Batch Write Back
            rangeH.formulas = formulasH;
            
            // Note: We sync in the parent function, but syncing here is safe too
            await context.sync();
            console.log("Gantt Logic: Averages Updated");

        } catch (error) {
            console.error("GanttLogic Error:", error);
        }
    }
};
