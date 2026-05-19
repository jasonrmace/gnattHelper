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
            const currentSheet = context.workbook.worksheets.getItem("GanttChart");

            // 1. LOCATE THE TRUE TEMPLATE SOURCE
            // We access the named range via the Workbook to find where it really lives.
            const namedRange = context.workbook.names.getItem("Level1Task").getRange();
            
            // We load the 'worksheet' property to ensure we copy from the correct sheet
            namedRange.load("worksheet, rowIndex");
            await context.sync();

            // 2. DEFINE SOURCE ROW (FROM THE TEMPLATE SHEET)
            const sourceSheet = namedRange.worksheet;
            const sourceRowIndex = namedRange.rowIndex;
            
            // Grab the Entire Row (A:XFD) from the SOURCE sheet
            const sourceRow = sourceSheet.getRange(`${sourceRowIndex + 1}:${sourceRowIndex + 1}`);

            // 3. FIND INSERTION POINT (ON GANTT CHART)
            const footerRange = currentSheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();
            
            const footerRowIndex = footerRange.rowIndex;

            // 4. INSERT & COPY
            const targetRow = currentSheet.getRange(`${footerRowIndex + 1}:${footerRowIndex + 1}`);
            targetRow.insert(Excel.InsertShiftDirection.down);
            
            // Re-target the new blank row
            const newRow = currentSheet.getRange(`${footerRowIndex + 1}:${footerRowIndex + 1}`);
            
            // Copy everything (Formulas + Formatting) from the Source Sheet
            newRow.copyFrom(sourceRow);

            // 5. UPDATE COLUMN B ONLY
            // We do not touch Column A, so the copied formula stays safe.
            const cellName = currentSheet.getCell(footerRowIndex, 1); 
            cellName.values = [[nameInput.value]];

            await context.sync();
            
            msg.innerText = "Success! Project created.";
            msg.className = "mt-4 text-sm text-center text-green-600";
            nameInput.value = "";
        });
    } catch (error) {
        console.error(error);
        if (error.message.includes("Level1Task")) {
             msg.innerText = "Error: Named Range 'Level1Task' not found.";
        } else {
             msg.innerText = "Error: " + error.message;
        }
        msg.className = "mt-4 text-sm text-center text-red-500";
    }
}
