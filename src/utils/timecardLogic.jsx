/* global Excel, Office */
import React from 'react';
import { Button } from 'react-bootstrap';
import { PDFDocument, PDFName } from 'pdf-lib';

export const TimecardLogic = {
    isProcessing: false,

    /**
     * Handles floating point precision issues (e.g. 7.1e-15) by using a tolerance
     */
    isErrorValue: (val) => Math.abs(parseFloat(val) || 0) > 0.01,

    /**
     * Manually constructs a digital signature field in the PDF structure.
     * This is used when the high-level library API is unavailable or 
     * fails to register the field correctly for viewers like Bluebeam.
     */
    createManualSignatureField: (pdfDoc, page, name, x, y, width, height) => {
        const { context } = pdfDoc;

        // 1. Create the Signature Widget Annotation (doubles as the field)
        const widgetDict = context.obj({
            Type: 'Annot',
            Subtype: 'Widget',
            FT: 'Sig',
            Rect: [x, y, x + width, y + height],
            T: context.objString(name),
            F: 4,      // Printable
            P: page.ref,
        });
        const widgetRef = context.register(widgetDict);

        // 2. Add annotation to the page
        let annots = context.lookup(page.node.get(PDFName.of('Annots')));
        if (!annots) {
            annots = context.obj([]);
            page.node.set(PDFName.of('Annots'), annots);
        }
        annots.push(widgetRef);

        // 3. Register field in document AcroForm
        const catalog = context.lookup(context.trailer.get(PDFName.of('Root')));
        let acroForm = context.lookup(catalog.get(PDFName.of('AcroForm')));
        if (!acroForm) {
            acroForm = context.obj({ Fields: [] });
            catalog.set(PDFName.of('AcroForm'), acroForm);
        }
        const fields = context.lookup(acroForm.get(PDFName.of('Fields'))) || context.obj([]);
        if (!acroForm.get(PDFName.of('Fields'))) {
            acroForm.set(PDFName.of('Fields'), fields);
        }
        fields.push(widgetRef);
        acroForm.set(PDFName.of('NeedAppearances'), context.boolean(true));
    },

    /**
     * Handles the submission process: 
     * 1. Changes tab color to Orange/Pending
     * 2. Generates PDF of current sheet
     * 3. Injects signature field via pdf-lib
     * 4. Triggers download/view
     */
    submitTimesheet: async (periodName) => {
        return await Excel.run(async (context) => {
            const workbook = context.workbook;
            const sheets = workbook.worksheets;
            const currentSheet = sheets.getItem(periodName);

            // 1. Set Tab Color to Orange (#FFC000)
            currentSheet.tabColor = "#FFC000";
            
            // 2. Prepare for PDF: Hide others to ensure single sheet export
            sheets.load("items");
            await context.sync();
            const visibilityStates = sheets.items.map(s => ({ name: s.name, visibility: s.visibility }));
            
            sheets.items.forEach(s => {
                if (s.name !== periodName) s.visibility = Excel.SheetVisibility.hidden;
            });
            await context.sync();

            // 3. Get PDF Bytes from Office JS
            const pdfBytes = await new Promise((resolve, reject) => {
                Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: 65536 }, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const file = result.value;
                        const slices = [];
                        let gotSlices = 0;
                        const getSlice = (index) => {
                            file.getSliceAsync(index, (sliceResult) => {
                                if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                                    slices[index] = sliceResult.value.data;
                                    gotSlices++;
                                    if (gotSlices === file.sliceCount) {
                                        file.closeAsync();
                                        const totalLength = slices.reduce((acc, s) => acc + s.length, 0);
                                        const combined = new Uint8Array(totalLength);
                                        let offset = 0;
                                        for (const slice of slices) { combined.set(slice, offset); offset += slice.length; }
                                        resolve(combined);
                                    } else { getSlice(gotSlices); }
                                } else { file.closeAsync(); reject(sliceResult.error); }
                            });
                        };
                        getSlice(0);
                    } else { reject(result.error); }
                });
            });

            // const signatureWidget = pdfDoc.context.obj({
            //     Type: 'Annot',
            //     Subtype: 'Widget',
            //     FT: 'Sig',
            //     Rect: [50, 50, 200, 100],
            //     T: pdfDoc.context.objString('employee_signature'),
            //     F: 4,
            //     P: page.ref,
            // })

            // 4. Restore original visibility
            sheets.items.forEach(s => {
                const original = visibilityStates.find(v => v.name === s.name);
                if (original) s.visibility = original.visibility;
            });
            await context.sync();

                        // 5. Add Form Signature Field (pdf-lib)
            let finalPdfBytes = pdfBytes;
            try {
                const pdfDoc = await PDFDocument.load(pdfBytes);
                const pages = pdfDoc.getPages();
                const firstPage = pages[0];
                const { height } = firstPage.getSize();

                const x = 2.9 * 72;              // 2.9 inches from left
                const yFromTop = 7.5 * 72;       // 7.5 inches from top
                const width = 3 * 72;            // 3 inches wide
                const heightBox = 36;            // 0.5 inch high
                const y = height - yFromTop - heightBox;

                // Call the manual builder
                TimecardLogic.createManualSignatureField(pdfDoc, firstPage, 'employee_signature', x, y, width, heightBox);

                finalPdfBytes = await pdfDoc.save();
            } catch (pdfErr) {

                console.warn("pdf-lib processing failed, providing raw PDF:", pdfErr);
            }

            // 6. Provide Download & Open View
            const blob = new Blob([finalPdfBytes], { type: 'application/pdf' });
            const url = URL.createObjectURL(blob);
            
            // Trigger browser download
            // This is the reliable way to "save" the file. 
            // Most browsers will save to 'Downloads' or prompt if configured.
            const a = document.createElement('a');
            a.href = url;
            a.download = `${periodName}_Timesheet.pdf`;
            a.click();

            if (window.GlobalToast) window.GlobalToast.success("Timesheet Submitted! PDF generated.");
            if (window.GlobalRefresh) window.GlobalRefresh();
            
            return true;
        });
    },

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