/* global Office, Excel */

// ==============================================================================
// IDENTITY LOGIC ENGINE (Phase 3: Full Name Support)
// Purpose: Manages "Who am I?" state using LocalStorage.
// Note: LocalStorage is shared across all Excel files using this add-in on this machine.
// CHANGE: Now exports a module instead of attaching to window.
// ==============================================================================

export const IdentityLogic = {
    // KEY for storage
    STORAGE_KEY: "barbizon_user_identity",

    // 1. GET CURRENT IDENTITY
    getIdentity: () => {
        try {
            // Updated: referencing the const directly, not window
            return localStorage.getItem(IdentityLogic.STORAGE_KEY);
        } catch (e) {
            console.warn("LocalStorage access denied:", e);
            return null;
        }
    },

    // 2. SET IDENTITY
    setIdentity: (name) => {
        try {
            localStorage.setItem(IdentityLogic.STORAGE_KEY, name);
            console.log(`Identity saved: ${name}`);
        } catch (e) {
            console.error("Failed to save identity:", e);
        }
    },

    // 3. CLEAR IDENTITY
    clearIdentity: () => {
        localStorage.removeItem(IdentityLogic.STORAGE_KEY);
    },

    // 4. FETCH TEAM LIST (Returns Objects: { first: "Chris", full: "Chris Smith" })
    fetchTeamMembers: async () => {
        try {
            return await Excel.run(async (context) => {
                const table = context.workbook.worksheets.getItem("Team").tables.getItem("Team");

                // 1. Get First Name Column (Required)
                const firstCol = table.columns.getItem("First Name").getDataBodyRange();
                firstCol.load("values");

                // 2. Get Last Name Column (Optional - safer check)
                const lastCol = table.columns.getItemOrNullObject("Last Name");
                lastCol.load("isNullObject");

                await context.sync();

                // 3. Load Last Name Data if column exists
                let lastNames = [];
                if (!lastCol.isNullObject) {
                    const lastBody = lastCol.getDataBodyRange();
                    lastBody.load("values");
                    await context.sync();
                    lastNames = lastBody.values;
                }

                // 4. Build the List
                const firstNames = firstCol.values;
                const members = [];
                const seen = new Set();

                for (let i = 0; i < firstNames.length; i++) {
                    const first = firstNames[i][0]?.toString().trim();

                    if (first && !seen.has(first)) {
                        // Get Last Name safely (if column exists and row has data)
                        const last = (lastNames.length > i && lastNames[i][0]) ? lastNames[i][0].toString().trim() : "";

                        members.push({
                            first: first, // Saved to LocalStorage (e.g., "Chris")
                            full: last ? `${first} ${last}` : first // Shown in Dropdown (e.g., "Chris Smith")
                        });
                        seen.add(first);
                    }
                }

                // Sort alphabetically by First Name
                return members.sort((a, b) => a.first.localeCompare(b.first));
            });
        } catch (error) {
            console.error("Error fetching team (Check 'First Name' column):", error);
            return [];
        }
    }
};
