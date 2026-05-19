Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // 1. SET VERSION NUMBER VISUALLY
        const vNum = "v7.0 (Stable)";
        const footer = document.getElementById("version-footer");
        if (footer) footer.innerText = vNum; 
        console.log("Loaded: " + vNum);

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

            // --- STEP 1: SAFE SEARCH (NO CRASHING) ---
            // We try the Sheet first. If it fails, we try the Workbook.
            // "getItemOrNullObject" returns a null object instead of throwing an error.
            
            let namedItem = sheet.names.getItemOrNullObject("Level1Task");
            await context.sync();

            if (namedItem.isNullObject) {
                // Not on sheet? Check Workbook.
                namedItem = context.workbook.names.getItemOrNullObject("Level1Task");
                await context.sync();
            }

            // If BOTH are null, then it really doesn't exist.
            if (namedItem.isNullObject) {
                throw new Error("Named Range 'Level1Task' not found. Check Name Manager.");
            }

            // --- STEP 2: GET ENTIRE ROW (Fixes A/B Formatting) ---
            // We grab the whole row so A aligns with A, and B aligns with B.
            sourceRow = namedItem.getRange().getEntireRow();

            // --- STEP 3: INSERT ---
            const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
            footerRange.load("rowIndex");
            await context.sync();
            
            const footerIndex = footerRange.rowIndex;
            const targetRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
            targetRow.insert(Excel.InsertShiftDirection.down);

            // --- STEP 4: COPY ALL (Fixes Formulas) ---
            const newRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
            newRow.copyFrom(sourceRow); // Copies Formulas + Formats

            // --- STEP 5: UPDATE NAME ---
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
