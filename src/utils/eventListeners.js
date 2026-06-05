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
// Only notifies about remote updates; does NOT run logic engines.
async function handleGanttChange(event) {
    if (event.source === Excel.EventSource.remote) {
        if (window.RefreshBadge) window.RefreshBadge();
    }

    // Structural changes (Insert/Delete) no longer require a full formatting reset
    // because rules are applied to a large buffer (Row 8-2000) and Excel 
    // handles range adjustments natively.
    return;
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
    }, 1500); // 1.5 second delay
}
