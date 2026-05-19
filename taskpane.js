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

            // 1. LOCATE TEMPLATE ROW
            // We find the row index of "Level1Task"
            const namedRange = sheet.names.getItem("Level1Task").getRange();
            namedRange.load("rowIndex");
            await context.sync();

            // Define the Source Row (The entire row A:XFD)
            const sourceRowIndex = namedRange.rowIndex;
            const sourceRow = sheet.getRange(`${sourceRowIndex + 1}:${sourceRowIndex + 1}`);

            // 2. FIND FOOTER
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();
            
            const footerRowIndex = footerRange.rowIndex;

            // 3. INSERT NEW ROW
            // Insert a blank row at the footer location, pushing footer down
            const insertRange = sheet.getRange(`${footerRowIndex + 1}:${footerRowIndex + 1}`);
            insertRange.insert(Excel.InsertShiftDirection.down);

            // 4. COPY EVERYTHING (Formulas + Formats)
            // We target the newly created blank row
            const newRow = sheet.getRange(`${footerRowIndex + 1}:${footerRowIndex + 1}`);
            
            // CRITICAL CHANGE: No second argument implies "Copy All" (Formulas, Formats, Values)
            newRow.copyFrom(sourceRow);

            // 5. UPDATE ONLY THE NAME (Column B)
            // We leave Column A alone so the copied formula does its work.
            const cellName = sheet.getCell(footerRowIndex, 1); // Index 1 = Column B
            cellName.values = [[nameInput.value]];

            await context.sync();
            
            msg.innerText = `Success! Project created.`;
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
