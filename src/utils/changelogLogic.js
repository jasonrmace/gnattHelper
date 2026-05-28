/* global Excel */
import { IdentityLogic } from './identityLogic';

export const ChangelogLogic = {
    // 1. LOG A NEW CHANGE
    logChange: async (context, changeDescription) => {
        try {
            // Try lowercase then capitalized to handle case sensitivity
            let table = context.workbook.tables.getItemOrNullObject("changelog");
            table.load("isNullObject");
            await context.sync();

            if (table.isNullObject) {
                table = context.workbook.tables.getItemOrNullObject("Changelog");
                table.load("isNullObject");
                await context.sync();
            }

            const user = IdentityLogic.getIdentity() || "Unknown";
            const now = new Date().toLocaleString();
            
            if (table.isNullObject) {
                console.warn("Changelog table not found. Skipping log.");
                return;
            }
            
            table.columns.load("items/name");
            await context.sync();

            // Build a dynamic map of column names to indices
            const colMap = {};
            table.columns.items.forEach((col, index) => { colMap[col.name] = index; });

            // Create row data based on dynamic positions
            // Structure: [Change, Who Did It, Who Has Seen, Date/Time]
            const rowData = new Array(table.columns.items.length).fill("");
            if (colMap["Change"] !== undefined) rowData[colMap["Change"]] = changeDescription;
            if (colMap["Who Did It"] !== undefined) rowData[colMap["Who Did It"]] = user;
            if (colMap["Who Has Seen"] !== undefined) rowData[colMap["Who Has Seen"]] = user;
            if (colMap["Date/Time"] !== undefined) rowData[colMap["Date/Time"]] = now;

            table.rows.add(null, [rowData]);
            await context.sync();
        } catch (error) {
            console.error("Log Change Error:", error);
        }
    },

    // 2. GET COUNT OF UNSEEN CHANGES
    getUnseenCount: async (context) => {
        try {
            const table = context.workbook.tables.getItemOrNullObject("changelog");
            const currentUser = IdentityLogic.getIdentity();
            if (!currentUser) return 0;

            const columns = table.columns;
            columns.load("items/name");
            const range = table.getDataBodyRange();
            range.load("values");
            await context.sync();

            const seenIdx = columns.items.findIndex(c => c.name === "Who Has Seen");
            if (seenIdx === -1) return 0;

            const values = range.values;
            let count = 0;
            for (let i = 0; i < values.length; i++) {
                const seenRaw = values[i][seenIdx] ? values[i][seenIdx].toString() : "";
                const seenList = seenRaw.split(",").map(s => s.trim()).filter(s => s !== "");
                if (!seenList.includes(currentUser)) count++;
            }
            return count;
        } catch (e) {
            return 0;
        }
    },

    // 3. GET ALL LOGS FOR THE UPDATES PAGE
    getAllLogs: async (context) => {
        try {
            const table = context.workbook.tables.getItemOrNullObject("changelog");
            const currentUser = IdentityLogic.getIdentity();
            const columns = table.columns;
            columns.load("items/name");
            const range = table.getDataBodyRange();
            range.load("values");
            await context.sync();

            const colMap = {};
            columns.items.forEach((col, index) => { colMap[col.name] = index; });

            return range.values.map(row => {
                const seenRaw = row[colMap["Who Has Seen"]] ? row[colMap["Who Has Seen"]].toString() : "";
                const seenList = seenRaw.split(",").map(s => s.trim());
                const author = row[colMap["Who Did It"]] || "Unknown";
                const desc = row[colMap["Change"]] || "";

                return {
                    text: `${author}: ${desc}`,
                    isNew: !seenList.includes(currentUser),
                    timestamp: row[colMap["Date/Time"]]
                };
            }).reverse(); // Newest first
        } catch (e) {
            return [];
        }
    },

    // 4. MARK ALL AS SEEN
    markAllAsSeen: async () => {
        try {
            await Excel.run(async (context) => {
                const table = context.workbook.tables.getItemOrNullObject("changelog");
                const currentUser = IdentityLogic.getIdentity();
                if (!currentUser) return;

                const columns = table.columns;
                columns.load("items/name");
                const range = table.getDataBodyRange();
                range.load("values");
                await context.sync();

                const seenIdx = columns.items.findIndex(c => c.name === "Who Has Seen");
                if (seenIdx === -1) return;

                const values = range.values;
                for (let i = 0; i < values.length; i++) {
                    const seenRaw = values[i][seenIdx] ? values[i][seenIdx].toString() : "";
                    const seenList = seenRaw.split(",").map(s => s.trim()).filter(s => s !== "");
                    
                    if (!seenList.includes(currentUser)) {
                        seenList.push(currentUser);
                        range.getCell(i, seenIdx).values = [[seenList.join(", ")]];
                    }
                }
                await context.sync();
            });
        } catch (error) {
            console.error("Error marking all seen:", error);
        }
    },

    // 5. GET UNSEEN ADMIN UPDATES (FOR STARTUP TOASTS)
    getUnseenAdminLogs: async (context) => {
        try {
            const table = context.workbook.tables.getItemOrNullObject("changelog");
            const currentUser = IdentityLogic.getIdentity();
            if (!currentUser) return [];

            const columns = table.columns;
            columns.load("items/name");
            const range = table.getDataBodyRange();
            range.load("values");
            await context.sync();

            const colMap = {};
            columns.items.forEach((col, index) => { colMap[col.name] = index; });

            const authorIdx = colMap["Who Did It"];
            const seenIdx = colMap["Who Has Seen"];
            const changeIdx = colMap["Change"];
            const timeIdx = colMap["Date/Time"];

            if (authorIdx === undefined || seenIdx === undefined) return [];

            const values = range.values;
            return values
                .filter(row => {
                    const author = row[authorIdx] || "";
                    const seenRaw = row[seenIdx] || "";
                    const seenList = seenRaw.split(",").map(s => s.trim());
                    return author === "Admin" && !seenList.includes(currentUser);
                })
                .map(row => ({
                    change: row[changeIdx],
                    timestamp: row[timeIdx]
                }));
        } catch (e) {
            return [];
        }
    },

    // (Legacy/Internal) MARK A SPECIFIC CHANGE AS SEEN
    markAsSeen: async (changeDescription, timestamp) => {
        try {
            await Excel.run(async (context) => {
                const table = context.workbook.tables.getItemOrNullObject("changelog");
                const currentUser = IdentityLogic.getIdentity();
                const columns = table.columns;
                columns.load("items/name");
                const range = table.getDataBodyRange();
                range.load("values");
                await context.sync();

                const colMap = {};
                columns.items.forEach((col, index) => { colMap[col.name] = index; });
                const seenIdx = colMap["Who Has Seen"];
                const changeIdx = colMap["Change"];
                const timeIdx = colMap["Date/Time"];

                const values = range.values;
                for (let i = 0; i < values.length; i++) {
                    // Match row by description and timestamp
                    if (values[i][changeIdx] === changeDescription && values[i][timeIdx] === timestamp) {
                        const seenRaw = values[i][seenIdx] ? values[i][seenIdx].toString() : "";
                        const seenList = seenRaw.split(",").map(s => s.trim()).filter(s => s !== "");
                        
                        if (seenList.includes(currentUser)) break;
                        
                        seenList.push(currentUser);
                        range.getCell(i, seenIdx).values = [[seenList.join(", ")]];
                        break;
                    }
                }
                await context.sync();
            });
        } catch (error) {
            console.error("Error marking change as seen:", error);
        }
    }
};