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
    // UPDATED VERSION
    const [version] = useState("v4.2.5");
    const [activeTab, setActiveTab] = useState("ProjectList");
    
    const [isValidFile, setIsValidFile] = useState(true);
    const [currentName, setCurrentName] = useState("");
    const [hudText, setHudText] = useState("Ready");
    const processingRef = useRef(false);

    // --- 1. INITIALIZATION & EVENTS ---
    useEffect(() => {
        Office.onReady(async (info) => {
            if (info.host === Office.HostType.Excel) {
                console.log("Office Ready");

                // A. File Validation
                const url = Office.context.document.url;
                const decodedUrl = decodeURI(url);
                const segments = decodedUrl.split('/');
                const fileName = segments[segments.length - 1];
                setCurrentName(fileName);

                if (!decodedUrl.includes(REQUIRED_FILENAME)) {
                     // setIsValidFile(false); // Disabled for testing
                }

                // B. Register Event Listeners
                Office.context.document.addHandlerAsync(
                    Office.EventType.DocumentSelectionChanged, 
                    handleSelectionChange
                );

                // C. AUTO-RUN VISUAL SYNC (One time on load)
                if (window.VisualLogic) {
                    await Excel.run(async (context) => {
                        await window.VisualLogic.refreshGridAlerts(context);
                    });
                }
            }
        });
        
        return () => {
            Office.context.document.removeHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                handleSelectionChange
            );
        };
    }, []);

    // --- 2. HUD LOGIC ---
    const handleSelectionChange = async () => {
        if (processingRef.current) return;
        processingRef.current = true;

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = context.workbook.getSelectedRange();
                range.load(["rowIndex", "rowCount"]);
                await context.sync();

                if (range.rowCount > 1 || range.rowIndex < 7) {
                    processingRef.current = false;
                    return;
                }

                const currentRowIndex = range.rowIndex;
                const idCell = sheet.getRangeByIndexes(currentRowIndex, 0, 1, 1);
                const nameCell = sheet.getRangeByIndexes(currentRowIndex, 1, 1, 1);
                idCell.load("text");
                nameCell.load("text");
                await context.sync();

                const currentID = idCell.text[0][0];
                const currentName = nameCell.text[0][0];
                let finalText = `Active Row: ${currentRowIndex + 1}`;

                if (currentID) {
                    const idNum = parseFloat(currentID);
                    if (Number.isInteger(idNum)) {
                        finalText = `PROJECT ${currentID}: ${currentName}`;
                    } else if (!isNaN(idNum)) {
                        const parentID = Math.floor(idNum).toString();
                        const foundRange = sheet.getRange("A:A").find(parentID, { completeMatch: true, matchCase: false });
                        const parentNameRange = foundRange.getOffsetRange(0, 1);
                        parentNameRange.load("text");
                        await context.sync(); // May fail if not found, good enough for now
                        const pName = parentNameRange.text[0][0];
                        finalText = `PROJECT ${parentID}: ${pName} (Task ${currentID})`;
                    }
                }

                setHudText(finalText);

                const shape = sheet.shapes.getItem("TaskHUD");
                shape.load("name"); 
                await context.sync().then(() => {
                     shape.textFrame.textRange.text = finalText;
                     return context.sync();
                }).catch(() => {});
            });
        } catch (error) {
            console.error("HUD Error:", error);
        } finally {
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

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
