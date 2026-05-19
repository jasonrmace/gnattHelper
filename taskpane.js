Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runButton").onclick = createNewProject;
    }
});

async function createNewProject() {
    const msg = document.getElementById("message");
    const nameInput = document.getElementById("projectName");
    
    if (!nameInput.value) {
        msg.innerText = "Please enter a name.";
        msg.className = "mt-4 text-sm text-center text-red-500";
        return;
    }

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("GanttChart");
            let sourceRow;
            let sourceRowIndex_Debug;

            // --- STEP 1: FIND THE TEMPLATE (Dual-Check) ---
            // We try to find "Level1Task" on the Sheet first.
            try {
                const sheetName = sheet.names.getItem("Level1Task");
                sourceRow = sheetName.getRange().getEntireRow();
                sourceRow.load("rowIndex");
                await context.sync();
            } catch (e) {
                // If Sheet-Level fails, try Workbook-Level
                try {
                    const wbName = context.workbook.names.getItem("Level1Task");
                    sourceRow = wbName.getRange().getEntireRow();
                    sourceRow.load("rowIndex");
                    await context.sync();
                } catch (e2) {
                    throw new Error("Could not find Named Range 'Level1Task' on Sheet OR Workbook.");
                }
            }

            sourceRowIndex_Debug = sourceRow.rowIndex + 1; // +1 for human readability

            // --- STEP 2: FIND INSERTION POINT ---
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();
            
            const footerIndex = footerRange.rowIndex;

            // --- STEP 3: INSERT & COPY ---
            // Create the blank space
            const targetRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
            targetRow.insert(Excel.InsertShiftDirection.down);
            
            // Retarget the new blank row
            const newRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);

            // FORCE COPY ALL (Formulas, Formats, Values, Validation)
            newRow.copyFrom(sourceRow, Excel.RangeCopyType.all);

            // --- STEP 4: UPDATE NAME ---
            // We ONLY touch Column B (Index 1). 
            // Column A is left strictly alone to preserve the formula copied from sourceRow.
            const cellName = sheet.getCell(footerIndex, 1);
            cellName.values = [[nameInput.value]];

            await context.sync();
            
            msg.innerText = `Success! Copied template from Row ${sourceRowIndex_Debug}.`;
            msg.className = "mt-4 text-sm text-center text-green-600";
            nameInput.value = "";
        });
    } catch (error) {
        console.error(error);
        msg.innerText = "Error: " + error.message;
        msg.className = "mt-4 text-sm text-center text-red-500";
    }
}
