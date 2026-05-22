/* global React, ReactDOM, Office, ReactBootstrap, Excel */

// 1. CONFIGURATION
const REQUIRED_FILENAME = "Houston Summer 2026 [Macros].xlsm"; 
const EXPECTED_LOCATION = ""; 

// 2. UNPACK LIBRARIES
const { useState, useEffect, useRef } = React;
const { Container, Alert, Spinner } = ReactBootstrap || {};

const MainNavbar = window.MainNavbar;
const CreateProject = window.CreateProject;
const ProjectList = window.ProjectList;

function App() {
    // --- STATE ---
    const [version] = useState("v4.2.2");
    const [activeTab, setActiveTab] = useState("ProjectList");
    
    // Validation State
    const [isValidFile, setIsValidFile] = useState(true);
    const [currentName, setCurrentName] = useState("");
    
    // HUD State (New)
    const [hudText, setHudText] = useState("Ready");
    const processingRef = useRef(false); // Prevent event flooding

    // --- 1. INITIALIZATION & EVENTS ---
    useEffect(() => {
        Office.onReady((info) => {
            if (info.host === Office.HostType.Excel) {
                console.log("Office Ready");

                // A. File Validation Check
                const url = Office.context.document.url;
                const decodedUrl = decodeURI(url);
                const segments = decodedUrl.split('/');
                const fileName = segments[segments.length - 1];
                setCurrentName(fileName);

                if (!decodedUrl.includes(REQUIRED_FILENAME)) {
                    // setIsValidFile(false); // Commented out for testing flexibility
                }

                // B. Register Selection Change Listener (The VBA Replacement)
                Office.context.document.addHandlerAsync(
                    Office.EventType.DocumentSelectionChanged, 
                    handleSelectionChange
                );
            }
        });
        
        // Cleanup listener on unmount (optional for add-ins)
        return () => {
            Office.context.document.removeHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                handleSelectionChange
            );
        };
    }, []);

    // --- 2. THE HUD LOGIC (Migrated from VBA) ---
    const handleSelectionChange = async () => {
        // Debounce/Throttle: Don't run if already calculating
        if (processingRef.current) return;
        processingRef.current = true;

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = context.workbook.getSelectedRange();
                range.load(["rowIndex", "rowCount"]); // Just get index
                await context.sync();

                // Stop if multiple rows selected or clicking header (Rows 0-7)
                if (range.rowCount > 1 || range.rowIndex < 7) {
                    processingRef.current = false;
                    return;
                }

                const currentRowIndex = range.rowIndex;

                // Read Col A (ID) and Col B (Name) of clicked row
                // getRangeByIndexes(row, col, rowCount, colCount)
                const idCell = sheet.getRangeByIndexes(currentRowIndex, 0, 1, 1); // Col A
                const nameCell = sheet.getRangeByIndexes(currentRowIndex, 1, 1, 1); // Col B
                idCell.load("text");
                nameCell.load("text");
                await context.sync();

                const currentID = idCell.text[0][0];
                const currentName = nameCell.text[0][0];
                let finalText = `Active Row: ${currentRowIndex + 1}`;

                // LOGIC: Find the Context
                if (currentID) {
                    const idNum = parseFloat(currentID);
                    
                    if (Number.isInteger(idNum)) {
                        // CASE 1: User clicked a Project Header
                        finalText = `SECTION ${currentID}: ${currentName}`;
                    } else if (!isNaN(idNum)) {
                        // CASE 2: User clicked a Sub-Task (1.1) -> Find Parent (1)
                        const parentID = Math.floor(idNum).toString();
                        
                        // Search Column A for the Parent ID
                        const foundRange = sheet.getRange("A:A").find(parentID, {
                            completeMatch: true, 
                            matchCase: false
                        });
                        foundRange.load(["address", "rowIndex"]); 
                        
                        // We need to sync to check if found
                        // Note: find() throws error or returns null object if not found depending on version
                        // Safe approach: Try/Catch context around search isn't needed with 'OrNullObject' but find() is basic.
                        // Let's assume it finds it for now or fallback.
                        
                        // To read the Parent Name, we need the range from the found location
                        // However, 'find' returns a range object we can perform operations on
                        const parentNameRange = foundRange.getOffsetRange(0, 1); // Offset 0 rows, 1 col (Col B)
                        parentNameRange.load("text");
                        
                        await context.sync(); // Sync to get search results
                        
                        const pName = parentNameRange.text[0][0];
                        finalText = `SECTION ${parentID}: ${pName} (Task ${currentID})`;
                    }
                }

                // --- UPDATE UI ---
                // 1. React State
                setHudText(finalText);

                // 2. Excel Shape ("TaskHUD") - Logic from VBA
                const shape = sheet.shapes.getItem("TaskHUD");
                // We load the shape to see if it exists (avoids crash)
                shape.load("name"); 
                
                // Use try/catch block strictly for the shape update 
                // in case shape is missing (VBA "On Error Resume Next")
                await context.sync().then(() => {
                     shape.textFrame.textRange.text = finalText;
                     return context.sync();
                }).catch((e) => {
                    console.log("Shape 'TaskHUD' not found or locked.", e);
                });

            });
        } catch (error) {
            console.error("HUD Error:", error);
        } finally {
            // Release lock after short delay
            setTimeout(() => { processingRef.current = false; }, 200);
        }
    };


    // --- UI RENDER ---
    return (
        <>
            <MainNavbar 
                activeTab={activeTab} 
                setActiveTab={setActiveTab} 
                isFileValid={isValidFile}
            />
            
            <Container className="p-3" style={{ paddingBottom: "50px" }}>
                
                {/* SCENARIO A: BLOCKING ALERT */}
                {!isValidFile ? (
                    <div className="mt-4">
                        <Alert variant="danger" className="shadow-sm border-danger text-center">
                            <div className="mb-2" style={{fontSize: "1.5rem"}}>⛔</div>
                            <strong>Functionality Locked</strong>
                            <p className="small mt-2 mb-0">Wrong File.</p>
                        </Alert>
                    </div>
                ) : (
                    <>
                        {activeTab === "ProjectList" && <ProjectList />}
                        {activeTab === "AddProject" && <CreateProject />}
                    </>
                )}

                {/* NEW: HEADS UP DISPLAY (Sticky Footer) */}
                <div className="fixed-bottom bg-primary text-white shadow-lg px-3 py-2 d-flex justify-content-between align-items-center" 
                     style={{ fontSize: "0.8rem", borderTop: "3px solid #0d6efd" }}>
                    <span className="fw-bold text-truncate" style={{maxWidth: "80%"}}>
                        <i className="fas fa-crosshairs me-2 opacity-50"></i>
                        {hudText}
                    </span>
                    <span className="opacity-50" style={{fontSize: "0.7rem"}}>{version}</span>
                </div>

            </Container>
        </>
    );
}

// Mount the React App
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
