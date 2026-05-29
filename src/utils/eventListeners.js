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
                const ganttSheet = context.workbook.worksheets.getItem("GanttChart");

                teamSheet.load("id, name");
                ganttSheet.load("id, name");

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
                ganttSheet.onChanged.add(handleGanttChange);
                console.log(`✅ Watching Gantt Sheet: ${ganttSheet.name}`);

                // 4. REGISTER 'VACATIONS' LISTENER (If on a different sheet)
                if (!vacTable.isNullObject) {
                    const vacSheet = vacTable.worksheet;
                    vacSheet.load("id, name");
                    await context.sync();

                    // Only add if it's actually a different sheet (to avoid double-firing)
                    if (vacSheet.id !== teamSheet.id) {
                        vacSheet.onChanged.add(handleMetadataChange);
                        console.log(`✅ Watching Sheet: ${vacSheet.name}`);
                    }
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
// Only notifies about remote updates; does NOT run logic engines.
async function handleGanttChange(event) {
    if (event.source === Excel.EventSource.remote) {
        if (window.RefreshBadge) window.RefreshBadge();
    }

      // EXIT: If this is a standard cell edit (RangeEdited), we ignore it.
    // We only trigger the heavy logic if a project/task row was inserted or deleted.
    if (event.changeType !== "RowInserted" && event.changeType !== "RowDeleted") {
        return;
    }

    console.log(`Structural change detected (${event.changeType}). Syncing Grid...`);

     // Reuse the debounce logic to handle bulk inserts smoothly
    if (debounceTimer) clearTimeout(debounceTimer);

     debounceTimer = setTimeout(async () => {
        await Excel.run(async (context) => {
            // Re-apply everything because indices have shifted
            await FormattingLogic.generateSmartRules(context);
            await VisualLogic.refreshGridAlerts(context);
        });
    }, 1000);
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
        console.log("🔄 Change detected. Syncing Grid...");

        await Excel.run(async (context) => {
            // STEP 1: Formatting (Heavy Lifting - Bars/Colors)
            console.log(">> Running Formatting Logic...");
            await FormattingLogic.generateSmartRules(context);

            // STEP 2: Visuals (Overlays - Counters/Popups)
            // Must run AFTER formatting because formatting clears the grid
            console.log(">> Running Visual Logic...");
            await VisualLogic.refreshGridAlerts(context);
        });
    }, 1500); // 1.5 second delay
}
