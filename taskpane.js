Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runButton").onclick = createNewProject;
    }
});

async function createNewProject() {
    console.log("Version 6.0 - Safe Fetch Mode"); 
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

            // --- STEP 1: SAFE SEARCH (The Fix) ---
            // We use 'getItemOrNullObject'. This prevents the "Resource doesn't exist" crash.
            
            // Check Sheet Level First
            let namedItem = sheet.names.getItemOrNullObject("Level1Task");
            await context.sync();

            // If not found on Sheet, check Workbook Level
            if (namedItem.isNullObject) {
                console.log("Not found on sheet, checking workbook...");
                namedItem = context.workbook.names.getItemOrNullObject("Level1Task");
                await context.sync();
            }

            // If STILL not found, stop safely.
            if (namedItem.isNullObject) {
                throw new Error("Could not find Named Range 'Level1Task'. Please check Spelling in Name Manager.");
            }

            // --- STEP 2: GET SOURCE ROW ---
            // We found it! Now get the row.
            const range = namedItem.getRange();
            sourceRow = range.getEntireRow();
            
            // Load the row index for the success message
            sourceRow.load("rowIndex"); 
            await context.sync();

            // --- STEP 3: FIND FOOTER ---
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();
            
            const footerIndex = footerRange.rowIndex;

            // --- STEP 4: INSERT & COPY (Formulas included) ---
            const targetRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
            targetRow.insert(Excel.InsertShiftDirection.down);
            
            const newRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
            
            // Copy EVERYTHING (Formulas, Formats, Values)
            newRow.copyFrom(sourceRow);

            // --- STEP 5: UPDATE NAME ---
            // Only update Column B (Index 1). Leave Col A alone.
            const cellName = sheet.getCell(footerIndex, 1);
            cellName.values = [[nameInput.value]];

            await context.sync();
            
            msg.innerText = `Success! Copied Template from Row ${sourceRow.rowIndex + 1}.`;
            msg.className = "mt-4 text-sm text-center text-green-600";
            nameInput.value = "";
        });
    } catch (error) {
        console.error(error);
        msg.innerText = "Error: " + error.message;
        msg.className = "mt-4 text-sm text-center text-red-500";
    }
}
