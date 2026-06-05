/* global Excel */
import { IdentityLogic } from './identityLogic';

export const ChangelogLogic = {
    // 1. LOG A NEW CHANGE
    logChange: async (context, changeDescription, authorOverride = null) => {
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

            const user = authorOverride || IdentityLogic.getIdentity() || "Unknown";
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
            
            // Admin check to determine filtering
            const adminUsers = ["Rob", "Kevin", "Rob Kreps", "Kevin Rittner"];
            const isAdmin = adminUsers.includes(currentUser);

            const columns = table.columns;
            columns.load("items/name");
            const range = table.getDataBodyRange();
            range.load("values");
            await context.sync();

            const seenIdx = columns.items.findIndex(c => c.name === "Who Has Seen");
            const changeIdx = columns.items.findIndex(c => c.name === "Change");
            if (seenIdx === -1 || changeIdx === -1) return 0;

            const values = range.values;
            let count = 0;
            for (let i = 0; i < values.length; i++) {
                const desc = values[i][changeIdx] ? values[i][changeIdx].toString() : "";
                const seenRaw = values[i][seenIdx] ? values[i][seenIdx].toString() : "";

                // Privacy Filter: If not admin, skip counting PTO changes for other users
                if (!isAdmin && desc.includes("PTO for") && !desc.includes(currentUser)) continue;

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
            
            const adminUsers = ["Rob", "Kevin", "Rob Kreps", "Kevin Rittner"];
            const isAdmin = adminUsers.includes(currentUser);

            const columns = table.columns;
            columns.load("items/name");
            const range = table.getDataBodyRange();
            range.load("values");
            await context.sync();

            const colMap = {};
            columns.items.forEach((col, index) => { colMap[col.name] = index; });

            const logs = range.values.map(row => {
                const seenRaw = row[colMap["Who Has Seen"]] ? row[colMap["Who Has Seen"]].toString() : "";
                const seenList = seenRaw.split(",").map(s => s.trim());
                const author = row[colMap["Who Did It"]] || "Unknown";
                const desc = row[colMap["Change"]] || "";

                return {
                    desc: desc, // Keep raw desc for filtering
                    text: `${author}: ${desc}`,
                    isNew: !seenList.includes(currentUser),
                    timestamp: row[colMap["Date/Time"]]
                };
            });

            // Privacy Filter: Remove PTO logs that don't belong to the current non-admin user
            return logs.filter(log => {
                if (isAdmin) return true;
                return !(log.desc.includes("PTO for") && !log.desc.includes(currentUser));
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
    },

    // 6. GET UNSEEN PTO UPDATES (FOR ROB & KEVIN)
    getUnseenPTOLogs: async (context) => {
        try {
            const table = context.workbook.tables.getItemOrNullObject("changelog");
            const teamTable = context.workbook.tables.getItemOrNullObject("Team");
            const currentUser = IdentityLogic.getIdentity();
            if (!currentUser) return [];

            // Load Team Manager mappings
            const teamRange = teamTable.getDataBodyRange();
            teamTable.columns.load("items/name");
            teamRange.load("values");
            
            const columns = table.columns;
            columns.load("items/name");
            const range = table.getDataBodyRange();
            range.load("values");
            await context.sync();

            const teamColMap = {};
            teamTable.columns.items.forEach((col, idx) => { teamColMap[col.name] = idx; });
            const managerMap = {};
            teamRange.values.forEach(row => {
                const fName = (row[teamColMap["First Name"]] || "").toString().trim();
                const mgr = (row[teamColMap["Manager"]] || "").toString().trim();
                if (fName) managerMap[fName] = mgr;
            });

            const colMap = {};
            columns.items.forEach((col, index) => { colMap[col.name] = index; });

            const seenIdx = colMap["Who Has Seen"];
            const changeIdx = colMap["Change"];
            const timeIdx = colMap["Date/Time"];
            const authorIdx = colMap["Who Did It"];

            const values = range.values;
            return values
                .filter(row => {
                    const desc = (row[changeIdx] || "").toString();
                    const seenRaw = (row[seenIdx] || "").toString();
                    const seenList = seenRaw.split(",").map(s => s.trim());
                    
                    if (!desc.includes("PTO for") || seenList.includes(currentUser)) return false;

                    // Extract the requester name from "Added PTO for [Name]: ..."
                    const requester = desc.split(":")[0].replace("Added PTO for ", "").trim();
                    const assignedManager = managerMap[requester];
                    
                    // Only show if the current user is the assigned manager
                    return assignedManager === currentUser;
                })
                .map(row => ({
                    change: row[changeIdx],
                    timestamp: row[timeIdx],
                    author: row[authorIdx]
                }));
        } catch (e) { return []; }
    }
};