Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runButton").onclick = createNewProject;
    }
});

async function createNewProject() {
    console.log("Version 4.0 - Workbook Scope Fix");
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

            // 1. FIX: LOOK IN WORKBOOK NAMES (Not Sheet Names)
            // This finds "Level1Task" even if it is scoped globally.
            const namedItem = context.workbook.names.getItem("Level1Task");
            
            // Get the range and extend it to the ENTIRE ROW (A:XFD)
            // This captures your formula in Col A and formats in Col B
            const sourceRange = namedItem.getRange();
            const sourceRow = sourceRange.getEntireRow();

            // 2. FIND FOOTER POSITION
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();
            
            const footerIndex = footerRange.rowIndex;

            // 3. INSERT BLANK ROW
            // Push the footer down to make space
            const targetRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
            targetRow.insert(Excel.InsertShiftDirection.down);

            // 4. PASTE TEMPLATE
            // Target the new blank row and copy everything from the source
            const newRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
            newRow.copyFrom(sourceRow);

            // 5. UPDATE NAME ONLY
            // We update Column B (Index 1). Column A is left untouched (keeping the formula).
            const cellName = sheet.getCell(footerIndex, 1);
            cellName.values = [[nameInput.value]];

            await context.sync();
            
            msg.innerText = "Success! Project created.";
            msg.className = "mt-4 text-sm text-center text-green-600";
            nameInput.value = "";
        });
    } catch (error) {
        console.error(error);
        // Improved Error Message
        if (error.code === "ItemNotFound") {
             msg.innerText = "Error: Named Range 'Level1Task' or Sheet 'GanttChart' not found.";
        } else {
             msg.innerText = "Error: " + error.message;
        }
        msg.className = "mt-4 text-sm text-center text-red-500";
    }
}
