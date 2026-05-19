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

            // 1. Find the Footer (Anchor)
            // We search Column A for the specific text
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", {
                completeMatch: false, // Partial match ok
                matchCase: false
            });
            
            footerRange.load("rowIndex");
            await context.sync(); // SYNC 1: Get the Row Index

            // 2. Calculate Next ID
            const footerRowIndex = footerRange.rowIndex;
            
            // Look at the cell immediately above the footer
            const lastIdCell = sheet.getCell(footerRowIndex - 1, 0); // Col A is 0
            lastIdCell.load("values");
            await context.sync(); // SYNC 2: Get the Value
            
            let newID = 1;
            const lastVal = lastIdCell.values[0][0];
            
            // Check if it's a number
            if (!isNaN(lastVal) && lastVal !== "") {
                newID = parseInt(lastVal) + 1;
            }

            // 3. Insert The Row
            // We select the row exactly where the footer is, and push down
            const insertRange = sheet.getRange(`A${footerRowIndex + 1}:A${footerRowIndex + 1}`);
            insertRange.getEntireRow().insert(Excel.InsertShiftDirection.down);

            // 4. Write Data
            // Note: Using JS allows us to write values directly. 
            // This will TRIGGER your existing VBA 'Worksheet_Change' logic automatically!
            const newRowIndex = footerRowIndex; // The footer moved down, so we write to the old footer slot
            
            const cellID = sheet.getCell(newRowIndex, 0); // A
            const cellName = sheet.getCell(newRowIndex, 1); // B
            
            cellID.values = [[newID]];
            cellName.values = [[nameInput.value]];

            // 5. Formatting (Optional - VBA usually handles this via Template, but we can force it)
            cellID.format.font.bold = true;
            
            await context.sync(); // SYNC 3: Commit Changes
            
            // Success Feedback
            msg.innerText = `Success! Project ${newID} created.`;
            msg.className = "mt-4 text-sm text-center text-green-600";
            nameInput.value = ""; // Clear input
        });
    } catch (error) {
        console.error(error);
        msg.innerText = "Error: " + error.message;
        msg.className = "mt-4 text-sm text-center text-red-500";
    }
}
