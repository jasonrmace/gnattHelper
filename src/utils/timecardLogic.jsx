/* global Excel */
import React from 'react';
import { Button } from 'react-bootstrap';

export const TimecardLogic = {
    isProcessing: false,

    /**
     * Automatically checks if a new pay period sheet needs to be generated.
     * Triggered on add-in load for Timecard files.
     */
    checkAndGenerateNextPeriod: async () => {
        if (TimecardLogic.isProcessing) return;
        TimecardLogic.isProcessing = true;
        let rolloverTriggered = false;

        try {
            return await Excel.run(async (context) => {
                const sheets = context.workbook.worksheets;
                
                // Load items to access by index safely
                sheets.load("items");
                await context.sync();
                const firstSheet = sheets.items[0];
                
                // 1. Read the end date (Friday) of the current period
                const endDateRange = firstSheet.getRange("I19");
                endDateRange.load(["values", "text"]);
                await context.sync();

                const excelEndDateValue = endDateRange.values[0][0];
                if (!excelEndDateValue) return;

                // Convert Excel serial date to JS Date for "Today" comparison
                const currentEndDate = new Date(Math.round((excelEndDateValue - 25569) * 86400 * 1000));
                const today = new Date();
                
                // Normalize dates to midnight for comparison
                const todayMidnight = new Date(today.getFullYear(), today.getMonth(), today.getDate());
                // We add the timezone offset back to the currentEndDate to treat it as local midnight
                const endMidnight = new Date(currentEndDate.getTime() + (currentEndDate.getTimezoneOffset() * 60000));

                // 2. LOGIC: Trigger ONLY if today is Friday AND it matches the current period's end date
                const isFriday = today.getDay() === 5;
                const isPeriodEnd = todayMidnight.getTime() === endMidnight.getTime();
                const shouldRollover = isFriday && isPeriodEnd;

                if (shouldRollover) {
                    console.log("Timecard: Period end detected. Checking for next sheet...");

                    // 3. Calculate New Dates using Serial Math (safest for Excel)
                    const nextStartSerial = excelEndDateValue + 1; // Saturday
                    const nextEndSerial = excelEndDateValue + 14;  // Next Friday

                    // Helper to convert serial to a Local Date object
                    const serialToLocalDate = (s) => {
                        const d = new Date(Math.round((s - 25569) * 86400 * 1000));
                        return new Date(d.getTime() + (d.getTimezoneOffset() * 60000));
                    };

                    rolloverTriggered = true;
                    return await TimecardLogic.createNewPeriod(serialToLocalDate(nextEndSerial), serialToLocalDate(nextStartSerial));
                }
                return false;
            });
        } catch (error) {
            console.error("Timecard Rollover Error:", error);
        } finally {
            if (window.GlobalLoader) window.GlobalLoader.hide();
            if (!rolloverTriggered) TimecardLogic.isProcessing = false;
        }
    },

    /**
     * Finds the latest end date among existing sheets to suggest the next one.
     */
    getLatestPeriodEndDate: async () => {
        return await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            let latest = null;
            for (const sheet of sheets.items) {
                const parts = sheet.name.split('.');
                if (parts.length === 3) {
                    const d = new Date(parts[2], parts[0] - 1, parts[1]);
                    if (!isNaN(d.getTime())) {
                        if (!latest || d > latest) latest = d;
                    }
                }
            }
            return latest;
        });
    },

    /**
     * Core function to copy the template and set dates.
     */
    createNewPeriod: async (endDate, startDateOverride = null) => {
        // Ensure we aren't already middle-of-creation
        TimecardLogic.isProcessing = true;

        try {
            if (window.GlobalLoader) window.GlobalLoader.show("Generating New Period...");

            return await Excel.run(async (context) => {
                const sheets = context.workbook.worksheets;
                
                // Load items to access the template sheet safely
                sheets.load("items");
                await context.sync();
                const templateSheet = sheets.items[0];

                // Format Name: MM.DD.YYYY
                const newSheetName = `${(endDate.getMonth() + 1).toString().padStart(2, '0')}.${endDate.getDate().toString().padStart(2, '0')}.${endDate.getFullYear()}`;

                // Check existence
                const existingSheet = sheets.getItemOrNullObject(newSheetName);
                existingSheet.load("name");
                await context.sync();

                if (!existingSheet.isNullObject) {
                    throw new Error(`A timesheet for ${newSheetName} already exists.`);
                }

                // Determine Start Date
                let startDate = startDateOverride;
                if (!startDate) {
                    startDate = new Date(endDate);
                    startDate.setDate(startDate.getDate() - 13); // Friday back to Saturday 2 weeks prior
                }

                // Perform Copy
                const newSheet = templateSheet.copy(Excel.WorksheetPositionType.before, templateSheet);
                newSheet.name = newSheetName;
                newSheet.tabColor = ""; // Clear "Submitted" green

                // Set C7 Start Date - Use a formatted string to avoid ISO/Timezone issues in cells
                const startCell = newSheet.getRange("C7");
                const dateStr = `${startDate.getMonth() + 1}/${startDate.getDate()}/${startDate.getFullYear()}`;
                startCell.values = [[dateStr]];
                startCell.numberFormat = [["m/d/yyyy"]];

                // Clear specific input ranges to make it a clean timesheet
                // Note: Individually cleared because getRange does not support comma-separated strings
                ["C8:I11", "C14:I16", "C20:I23", "C26:I28"].forEach(address => {
                    newSheet.getRange(address).clear(Excel.ClearApplyTo.contents);
                });
                
                newSheet.activate();
                await context.sync();

                if (window.GlobalToast) {
                     window.GlobalToast.success(
                        <span>
                            New Timesheet Created. 
                            <Button
                                variant="link" 
                                onClick={() => window.GlobalRefresh && window.GlobalRefresh()} 
                                className="p-0 m-0 text-decoration-none"
                            >Refresh </Button> 
                            available timesheets to view the new record.
                        </span>
                    );
                }
                return true;
            });
        } catch (error) {
            console.error("Create Period Error:", error);
            if (window.GlobalToast) {
                window.GlobalToast.error(error.message || "Failed to create timesheet.");
            }
            return false;
        } finally {
            if (window.GlobalLoader) window.GlobalLoader.hide();
            TimecardLogic.isProcessing = false;
        }
    }
};