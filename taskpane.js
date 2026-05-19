/* global React, ReactDOM, Office, Excel */

const { useState, useEffect } = React;

function App() {
    // --- STATE MANAGEMENT ---
    const [projectName, setProjectName] = useState("");
    const [status, setStatus] = useState({ msg: "Ready.", type: "text-muted" });
    const [isLoading, setIsLoading] = useState(false);
    const [version] = useState("React Split v2.0"); // Verify this in the UI

    // --- INITIALIZATION ---
    useEffect(() => {
        Office.onReady((info) => {
            if (info.host === Office.HostType.Excel) {
                console.log("Office Ready - React Loaded");
            }
        });
    }, []);

    // --- CORE LOGIC (The Safe Search + Copy Row) ---
    const handleCreate = async () => {
        if (!projectName) {
            setStatus({ msg: "Please enter a project name.", type: "text-danger" });
            return;
        }

        setIsLoading(true);
        setStatus({ msg: "Processing...", type: "text-primary" });

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");

                // 1. Safe Search for Template (Sheet -> Workbook)
                let namedItem = sheet.names.getItemOrNullObject("Level1Task");
                await context.sync();

                if (namedItem.isNullObject) {
                    namedItem = context.workbook.names.getItemOrNullObject("Level1Task");
                    await context.sync();
                }

                if (namedItem.isNullObject) {
                    throw new Error("Named Range 'Level1Task' not found. Check Name Manager.");
                }

                // 2. Get Source Row (Entire Row for formatting safety)
                const sourceRow = namedItem.getRange().getEntireRow();
                sourceRow.load("rowIndex"); 
                await context.sync();

                // 3. Find Footer
                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                footerRange.load("rowIndex");
                await context.sync();
                
                const footerIndex = footerRange.rowIndex;

                // 4. Insert & Copy
                // Insert Down
                const targetRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
                targetRow.insert(Excel.InsertShiftDirection.down);
                
                // Copy All (Formulas, Formats, Validation)
                const newRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
                newRow.copyFrom(sourceRow, Excel.RangeCopyType.all);

                // 5. Update Name (Column B only)
                const cellName = sheet.getCell(footerIndex, 1);
                cellName.values = [[projectName]];

                await context.sync();
                
                // Success Feedback
                setStatus({ 
                    msg: `Success! Added '${projectName}' (from Row ${sourceRow.rowIndex + 1}).`, 
                    type: "text-success" 
                });
                setProjectName(""); // Clear input
            });
        } catch (error) {
            console.error(error);
            setStatus({ msg: "Error: " + error.message, type: "text-danger" });
        } finally {
            setIsLoading(false);
        }
    };

    // --- THE UI RENDER ---
    return (
        <div className="container-fluid p-3">
            
            <div className="border-bottom border-primary pb-2 mb-3">
                <h5 className="text-primary fw-bold m-0">Gantt Manager</h5>
                <small className="text-muted" style={{fontSize: "0.7rem"}}>React Edition</small>
            </div>

            <div className="mb-3">
                <label className="form-label fw-bold small text-uppercase">New Project Name</label>
                <input 
                    type="text" 
                    className="form-control" 
                    placeholder="Enter name..."
                    value={projectName}
                    onChange={(e) => setProjectName(e.target.value)}
                    disabled={isLoading}
                />
            </div>

            <div className="d-grid gap-2">
                <button 
                    className="btn btn-primary" 
                    onClick={handleCreate}
                    disabled={isLoading}
                >
                    {isLoading ? (
                        <span><span className="spinner-border spinner-border-sm me-2"></span>Working...</span>
                    ) : (
                        "Create Project"
                    )}
                </button>
            </div>

            {/* Status Message Area */}
            <div className={`mt-3 text-center small fw-bold ${status.type}`}>
                {status.msg}
            </div>

            {/* Version Footer */}
            <div className="fixed-bottom p-1 text-end text-muted pe-3" style={{fontSize: "0.6rem", backgroundColor: "#f1f1f1"}}>
                {version}
            </div>
        </div>
    );
}

// Mount the React App
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
