Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("runButton").onclick = createNewProject;
    }
});

async function createNewProject() {
    const msg = document.getElementById("message");
    const nameInput = document.getElementById("projectName");
    
    if (!nameInput.value) return;

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("GanttChart");

            // 1. Find Footer
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();

            // 2. Calculate ID
            const footerRowIndex = footerRange.rowIndex;
            const lastIdCell = sheet.getCell(footerRowIndex - 1, 0);
            lastIdCell.load("values");
            await context.sync();
            
            let newID = 1;
            const lastVal = lastIdCell.values[0][0];
            if (!isNaN(lastVal) && lastVal !== "") newID = parseInt(lastVal) + 1;

            // 3. Insert & Write
            const insertRange = sheet.getRange(`A${footerRowIndex + 1}:A${footerRowIndex + 1}`);
            insertRange.getEntireRow().insert(Excel.InsertShiftDirection.down);

            sheet.getCell(footerRowIndex, 0).values = [[newID]];
            sheet.getCell(footerRowIndex, 1).values = [[nameInput.value]];
            sheet.getCell(footerRowIndex, 0).format.font.bold = true;

            await context.sync();
            
            msg.innerText = `Created Project ${newID}`;
            msg.className = "mt-4 text-sm text-green-600";
            nameInput.value = "";
        });
    } catch (error) {
        msg.innerText = "Error: " + error.message;
        msg.className = "mt-4 text-sm text-red-500";
    }
}
