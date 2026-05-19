Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runButton").onclick = createNewProject;
    }
});

async function createNewProject() {
    console.log("Version 3.0 - Entire Row Copy"); // Look for this in the console
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

            // 1. IDENTIFY TEMPLATE ROW
            // We get the named range, then expand it to the ENTIRE ROW (A:XFD)
            // This ensures we grab the Formula in Col A and Formats in Col B.
            const namedRange = sheet.names.getItem("Level1Task").getRange();
            const sourceRow = namedRange.getEntireRow();

            // 2. FIND FOOTER POSITION
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();
            
            const footerIndex = footerRange.rowIndex;

            // 3. INSERT BLANK ROW
            // We target the footer row and insert "Down". 
            // The footer moves to footerIndex + 1. The new blank row is at footerIndex.
            const targetRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
            targetRow.insert(Excel.InsertShiftDirection.down);

            // 4. PASTE TEMPLATE (ALL)
            // We copy the source row onto the new blank row.
            // copyFrom without arguments copies Values, Formulas, and Formats.
            const newRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
            newRow.copyFrom(sourceRow);

            // 5. UPDATE NAME ONLY
            // We ONLY touch Column B (Index 1). 
            // Column A is left alone, so the copied formula remains valid.
            const cellName = sheet.getCell(footerIndex, 1);
            cellName.values = [[nameInput.value]];

            await context.sync();
            
            msg.innerText = "Success! Project created.";
            msg.className = "mt-4 text-sm text-center text-green-600";
            nameInput.value = "";
        });
    } catch (error) {
        console.error(error);
        msg.innerText = "Error: " + error.message;
        msg.className = "mt-4 text-sm text-center text-red-500";
    }
}
