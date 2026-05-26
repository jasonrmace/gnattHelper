/* global Excel */

// ==============================================================================
// EVENT LISTENERS (The "Master Watchdog")
// Purpose: Watches 'Team' AND 'Vacations' tables. Triggers full grid sync.
// Logic:
// 1. Detects where the tables live.
// 2. Adds listeners to those sheets.
// 3. On Change: Runs Formatting (Bars) -> Then Visuals (Counters).
// ==============================================================================

window.EventListeners = {
    isRegistered: false,

    register: async () => {
        if (window.EventListeners.isRegistered) return;

        try {
            await Excel.run(async (context) => {
                // 1. IDENTIFY SHEETS
                // We assume "Team" sheet exists.
                // const teamSheet = context.workbook.worksheets.getItem("Team");
                const teamSheet = context.workbook.tables.getItemOrNullObject("Team");
                teamSheet.load("id, name");

                // We try to find the "Vacations" table to see where it lives.
                const vacTable = context.workbook.tables.getItemOrNullObject("Vacations");
                vacTable.load("worksheet");
                
                await context.sync();

                // 2. REGISTER 'TEAM' LISTENER
                // This handles Name/Color changes
                teamSheet.onChanged.add(handleSheetChange);
                console.log(`✅ Watching Sheet: ${teamSheet.name}`);

                // 3. REGISTER 'VACATIONS' LISTENER (If on a different sheet)
                if (!vacTable.isNullObject) {
                    const vacSheet = vacTable.worksheet;
                    vacSheet.load("id, name");
                    await context.sync();

                    // Only add if it's actually a different sheet (to avoid double-firing)
                    if (vacSheet.id !== teamSheet.id) {
                        vacSheet.onChanged.add(handleSheetChange);
                        console.log(`✅ Watching Sheet: ${vacSheet.name}`);
                    }
                } else {
                    console.log("⚠️ 'Vacations' table not found. Only watching 'Team' sheet.");
                }

                window.EventListeners.isRegistered = true;
            });
        } catch (error) {
            console.error("Failed to register listeners:", error);
        }
    }
};

// --- THE HANDLER (Debounced Global Refresh) ---
let debounceTimer = null;

async function handleSheetChange(event) {
    // Ignore changes made by the Add-in itself
    if (event.source === Excel.EventSource.remote) return;

    // Debounce: Wait 1.5 seconds after the last keystroke before running
    // This prevents the script from freezing Excel while you are typing
    if (debounceTimer) clearTimeout(debounceTimer);

    debounceTimer = setTimeout(async () => {
        console.log("🔄 Change detected. Syncing Grid...");

        await Excel.run(async (context) => {
            // STEP 1: Formatting (Heavy Lifting - Bars/Colors)
            if (window.FormattingLogic) {
                console.log(">> Running Formatting Logic...");
                await window.FormattingLogic.generateSmartRules(context);
            }

            // STEP 2: Visuals (Overlays - Counters/Popups)
            // Must run AFTER formatting because formatting clears the grid
            if (window.VisualLogic) {
                console.log(">> Running Visual Logic...");
                await window.VisualLogic.refreshGridAlerts(context);
            }
        });
    }, 1500); // 1.5 second delay
}
