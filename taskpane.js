/* global React, ReactDOM, Office, ReactBootstrap, Excel */

// 1. CONFIGURATION
const REQUIRED_FILENAME = "Houston Summer 2026 [Macros].xlsm";

// 2. UNPACK LIBRARIES
const { useState, useEffect, useRef } = React;
const { Container, Alert, Spinner } = ReactBootstrap || {};
const MainNavbar = window.MainNavbar;
const CreateProject = window.CreateProject;
const ProjectList = window.ProjectList;

// --- GLOBAL LOADING OVERLAY ---
const LoadingOverlay = ({ isVisible, message }) => {
    if (!isVisible) return null;
    
    return (
        <div style={{
            position: 'fixed', top: 0, left: 0, width: '100%', height: '100%',
            backgroundColor: 'rgba(255, 255, 255, 0.9)', zIndex: 9999,
            display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center'
        }}>
            <div className="text-primary mb-3">
                <i className="fas fa-spinner fa-spin fa-3x"></i>
            </div>
            <h5 className="text-dark fw-bold">{message}</h5>
            <p className="text-muted small">Please wait while Excel updates...</p>
        </div>
    );
};

function App() {
    // --- STATE ---
    const [version] = useState("v4.4.1");
    const [activeTab, setActiveTab] = useState("ProjectList");
    const [isValidFile, setIsValidFile] = useState(true);
    const [currentName, setCurrentName] = useState("");
    const [hudText, setHudText] = useState("Ready");
    
    // LOADER STATE
    const [isLoading, setIsLoading] = useState(false);
    const [loadingMsg, setLoadingMsg] = useState("Processing...");
    
    const processingRef = useRef(false);

    // --- EXPOSE LOADER TO WINDOW ---
    useEffect(() => {
        window.GlobalLoader = {
            show: (msg = "Updating...") => {
                setLoadingMsg(msg);
                setIsLoading(true);
            },
            hide: () => {
                setIsLoading(false);
            }
        };
    }, []);

    // --- INITIALIZATION & EVENTS ---
    useEffect(() => {
        Office.onReady(async (info) => {
            if (info.host === Office.HostType.Excel) {
                console.log("Office Ready");

                const url = Office.context.document.url;
                const decodedUrl = decodeURI(url);
                const segments = decodedUrl.split('/');
                setCurrentName(segments[segments.length - 1]);

                // A. Register HUD Selection Listener
                Office.context.document.addHandlerAsync(
                    Office.EventType.DocumentSelectionChanged,
                    handleSelectionChange
                );

                // B. REGISTER WATCHDOG (Silent)
                // This turns on the "Ears" but does NOT run the scripts yet.
                // Scripts will only run when a 'SheetChanged' event fires.
                if (window.EventListeners) {
                    await window.EventListeners.register();
                }
                
                // REMOVED: Section D (Initial Full Sync)
                // The add-on now loads silently without touching the grid.
            }
        });

        return () => {
            Office.context.document.removeHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                handleSelectionChange
            );
        };
    }, []);

    // --- HUD LOGIC (Unchanged) ---
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
                idCell.load("text"); nameCell.load("text"); 
                await context.sync();

                const currentID = idCell.text[0][0];
                const currentName = nameCell.text[0][0];
                let finalText = `Active Row: ${currentRowIndex + 1}`;

                if (currentID) {
                    const idNum = parseFloat(currentID);
                    if (Number.isInteger(idNum)) {
                        finalText = `PROJECT ${currentID}: ${currentName}`;
                    } else if (!isNaN(idNum)) {
                        finalText = `PROJECT ${currentID}: Task`; 
                    }
                }
                setHudText(finalText);
                
                // Update Shape
                const shape = sheet.shapes.getItem("TaskHUD");
                shape.load("name"); // Check existence
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
        <div className="d-flex flex-column vh-100 bg-white position-relative">
            
            {/* GLOBAL LOADER */}
            <LoadingOverlay isVisible={isLoading} message={loadingMsg} />

            {/* 1. HEADER */}
            <div className="flex-shrink-0">
                <MainNavbar activeTab={activeTab} setActiveTab={setActiveTab} isFileValid={isValidFile} />
            </div>

            {/* 2. BODY */}
            <div className="flex-grow-1 overflow-auto p-3">
                {!isValidFile ? (
                    <div className="mt-4">
                        <Alert variant="danger" className="shadow-sm border-danger text-center">
                            <strong>Functionality Locked</strong>
                        </Alert>
                    </div>
                ) : (
                    <>
                        {activeTab === "ProjectList" && <ProjectList />}
                        {activeTab === "AddProject" && <CreateProject />}
                    </>
                )}
            </div>

            {/* 3. FOOTER */}
            <div className="bg-primary text-white shadow-lg px-3 py-2 d-flex justify-content-between align-items-center flex-shrink-0" 
                 style={{ fontSize: "0.8rem", borderTop: "3px solid #0d6efd", zIndex: 1030 }}>
                <span className="fw-bold text-truncate" style={{maxWidth: "80%"}}>
                    <i className="fas fa-crosshairs me-2 opacity-50"></i> {hudText}
                </span>
                <span className="opacity-50" style={{fontSize: "0.7rem"}}>{version}</span>
            </div>
        </div>
    );
}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
