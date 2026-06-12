/* global Excel */
import { FormattingLogic } from './formattingLogic_v2';
import { VisualLogic } from './visualLogic';
import { ChangelogLogic } from './changelogLogic';

// ==============================================================================
// EVENT LISTENERS (The "Master Watchdog")
// Purpose: Watches 'Team' AND 'Vacations' tables. Triggers full grid sync.
// Logic:
// 1. Detects where the tables live.
// 2. Adds listeners to those sheets.
// 3. On Change: Runs Formatting (Bars) -> Then Visuals (Counters).
// ==============================================================================

const GANTT_SHEETS = ["Houston", "Dallas"];

export const EventListeners = {
    isRegistered: false,

    register: async () => {
        if (EventListeners.isRegistered) return;

        try {
            await Excel.run(async (context) => {
                // 1. IDENTIFY SHEETS
                // We assume "Team" sheet exists.
                const teamTable = context.workbook.tables.getItemOrNullObject("Team");
                const teamSheet = context.workbook.worksheets.getItem("Team");

                teamSheet.load("id, name");

                // We try to find the "Vacations" table to see where it lives.
                const vacTable = context.workbook.tables.getItemOrNullObject("Vacations");
                vacTable.load("worksheet");
                
                await context.sync();

                // 2. REGISTER 'TEAM' LISTENER (Table-specific for precision)
                if (!teamTable.isNullObject) {
                    teamTable.onChanged.add(handleMetadataChange);
                    console.log(`✅ Watching Team Table`);
                } else {
                    // Fallback to sheet listener if table isn't named yet
                    teamSheet.onChanged.add(handleMetadataChange);
                    console.log(`✅ Watching Team Sheet: ${teamSheet.name}`);
                }

                // 3. REGISTER 'GANTT' LISTENER
                for (const name of GANTT_SHEETS) {
                    const sheet = context.workbook.worksheets.getItem(name);
                    sheet.onChanged.add(handleGanttChange);
                    console.log(`✅ Watching Gantt Sheet: ${name}`);
                }

                // 4. REGISTER 'VACATIONS' LISTENER (If on a different sheet)
                if (!vacTable.isNullObject) {
                    // Watch the table specifically, not the whole sheet, 
                    // to avoid triggering on Gantt task edits.
                    vacTable.onChanged.add(handleMetadataChange);
                    console.log(`✅ Watching Vacations Table`);
                } else {
                    console.log("⚠️ 'Vacations' table not found. Only watching 'Team' sheet.");
                }

                EventListeners.isRegistered = true;
            });
        } catch (error) {
            console.error("Failed to register listeners:", error);
        }
    },

    checkMissedChanges: async (context) => {
        return await ChangelogLogic.getUnseenCount(context);
    }
};

// --- HANDLER 1: GANTT CHANGES (Lightweight) ---
// Handles remote updates (badge/loader for co-author formatting) AND local changes (triggers logic engines).
let ganttDebounceTimer = null; // New debounce timer for Gantt changes
let isProcessingGantt = false; // Flag to prevent recursive event loops

async function handleGanttChange(event) {
    if (event.source === Excel.EventSource.remote) {
        if (window.RefreshBadge) window.RefreshBadge();

        // Check if a co-author is performing a heavy formatting update
        await Excel.run(async (context) => {
            const statusRange = context.workbook.names.getItemOrNullObject("GlobalFormattingStatus").getRangeOrNullObject();
            statusRange.load(["isNullObject", "values"]);
            await context.sync();

            if (!statusRange.isNullObject) {
                const status = statusRange.values[0][0];
                if (status === "IN_PROGRESS") {
                    if (window.GlobalLoader) window.GlobalLoader.show("A co-author is updating worksheet formatting...");
                } else {
                    if (window.GlobalLoader) window.GlobalLoader.hide();
                }
            }
        });
    } else if (event.source === Excel.EventSource.local) {
        // If the add-in is already processing a change, ignore events 
        // triggered by our own writes to prevent an infinite loop.
        if (isProcessingGantt) return;

        // Handle local changes to Gantt sheets
        // Debounce: Wait 1.5 seconds after the last keystroke before running
        if (ganttDebounceTimer) clearTimeout(ganttDebounceTimer);

        ganttDebounceTimer = setTimeout(async () => {
            let message = "Updating Gantt chart...";
            let sheetName = "";
            let rowNumber = "";

            try {
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getItem(event.worksheetId);
                    sheet.load("name");
                    await context.sync();
                    sheetName = sheet.name;

                    // Extract row number from address (e.g., "Sheet1!A1:C3" -> 1)
                    // This is a best-effort attempt; event.address might be complex.
                    const addressParts = event.address.split('!');
                    const rangePart = addressParts.length > 1 ? addressParts[1] : addressParts[0];
                    const rowMatch = rangePart.match(/\d+/);
                    if (rowMatch) {
                        rowNumber = rowMatch[0];
                    }

                    if (event.changeType === "CellChanged" || event.changeType === "RangeEdited") {
                        message = `Editing Task Row ${rowNumber} in ${sheetName}...`;
                    } else if (event.changeType === "RowInserted") {
                        message = `Adding Task Row(s) in ${sheetName}...`;
                    } else if (event.changeType === "RowDeleted") { // This will still trigger a full refresh
                        message = `Deleting Task Row(s) in ${sheetName}...`;
                    } else {
                        message = `Updating ${sheetName} sheet...`;
                    }
                    
                    if (window.GlobalLoader) window.GlobalLoader.show(message);

                    isProcessingGantt = true; // Block incoming events
                    console.log("🔄 Syncing Grid...");

                    // STEP 1: Formatting
                    if ((event.changeType === "CellChanged" || event.changeType === "RangeEdited") && rowNumber) {
                        console.log(`>> Running Surgical Formatting for ${sheetName} Row ${rowNumber}...`);
                        const metadata = await FormattingLogic.fetchFormattingMetadata(context);
                        // Apply rules only to the modified row (index is 0-based, so rowNumber - 1)
                        await FormattingLogic.applyRulesToRange(context, sheetName, parseInt(rowNumber) - 1, 1, 150, metadata);
                    } else {
                        console.log(`>> Running Full Formatting Logic for ${sheetName} due to structural change...`);
                        await FormattingLogic.generateSmartRules(context, sheetName);
                    }

                    // STEP 2: Visuals (Overlays - Counters/Popups)
                    // Only run for additions as requested; edits and deletes handle their own row logic.
                    if (event.changeType === "RowInserted") {
                        console.log(`>> Running Visual Logic for ${sheetName} due to row addition...`);
                        await VisualLogic.refreshGridAlerts(context, sheetName);
                    }
                });
            } catch (error) {
                console.error("Error processing local Gantt change:", error);
            } finally {
                isProcessingGantt = false; // Allow events again
                if (window.GlobalLoader) window.GlobalLoader.hide();
            }
        }, 1500); // 1.5 second delay
    }
}

// --- HANDLER 2: METADATA CHANGES (Debounced Global Refresh) ---
// Runs when Team or Vacation data changes.
let debounceTimer = null;

async function handleMetadataChange(event) {
    // Detect if this change came from another user (Remote)
    if (event.source === Excel.EventSource.remote) {
        if (window.GlobalToast) {
            window.GlobalToast.info("Team/Vacation data updated remotely. Syncing view...");
        }
    }

    // Debounce: Wait 1.5 seconds after the last keystroke before running
    // This prevents the script from freezing Excel while you are typing
    if (debounceTimer) clearTimeout(debounceTimer);

    debounceTimer = setTimeout(async () => {
        try {
            console.log("🔄 Change detected. Syncing Grid...");
            // Show a generic loader message for metadata changes,
            // which will be updated by FormattingLogic's more specific messages.
            if (window.GlobalLoader) window.GlobalLoader.show("Processing metadata changes...");

            await Excel.run(async (context) => {
                // STEP 1: Formatting (Heavy Lifting - Bars/Colors)
                console.log(">> Running Formatting Logic...");
                for (const name of GANTT_SHEETS) {
                    await FormattingLogic.generateSmartRules(context, name);
                }

                // STEP 2: Visuals (Overlays - Counters/Popups)
                // Must run AFTER formatting because formatting clears the grid
                console.log(">> Running Visual Logic...");
                for (const name of GANTT_SHEETS) {
                    await VisualLogic.refreshGridAlerts(context, name);
                }
            });
        } catch (error) {
            console.error("Error during metadata change processing:", error);
        } finally {
            if (window.GlobalLoader) window.GlobalLoader.hide();
        }
    }, 1500); // 1.5 second delay
}
