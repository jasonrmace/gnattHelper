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

            // 1. GET THE TEMPLATE (Simple Sheet-Level Fetch)
            const templateRange = sheet.names.getItem("Level1Task").getRange();

            // 2. Find Footer "DO NOT DELETE"
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync(); // Sync to get row index

            // 3. Calculate New ID
            const footerRowIndex = footerRange.rowIndex;
            const lastIdCell = sheet.getCell(footerRowIndex - 1, 0);
            lastIdCell.load("values");
            await context.sync(); 
            
            let newID = 1;
            const lastVal = lastIdCell.values[0][0];
            if (!isNaN(lastVal) && lastVal !== "") newID = parseInt(lastVal) + 1;

            // 4. Insert the New Row
            const newRowRange = sheet.getRange(`A${footerRowIndex + 1}:XFD${footerRowIndex + 1}`);
            newRowRange.insert(Excel.InsertShiftDirection.down);

            // 5. APPLY FORMATTING
            // This copies the formats from the template to the new blank row
            const insertedRow = sheet.getRange(`A${footerRowIndex + 1}:XFD${footerRowIndex + 1}`);
            insertedRow.copyFrom(templateRange, Excel.RangeCopyType.formats);

            // 6. Write Data (ID and Name)
            const cellId = sheet.getCell(footerRowIndex, 0);
            const cellName = sheet.getCell(footerRowIndex, 1);
            
            cellId.values = [[newID]];
            cellName.values = [[nameInput.value]];

            await context.sync();
            
            msg.innerText = `Success! Project ${newID} created.`;
            msg.className = "mt-4 text-sm text-center text-green-600";
            nameInput.value = "";
        });
    } catch (error) {
        console.error(error);
        if (error.message.includes("Level1Task")) {
             msg.innerText = "Error: Named Range 'Level1Task' not found on this sheet.";
        } else {
             msg.innerText = "Error: " + error.message;
        }
        msg.className = "mt-4 text-sm text-center text-red-500";
    }
}
