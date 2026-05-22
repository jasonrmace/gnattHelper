/* global React, ReactDOM, Office, ReactBootstrap */

// 1. CONFIGURATION
// The EXACT filename required to run the app
const REQUIRED_FILENAME = "Houston Summer 2026 [Macros].xlsm";

// 2. UNPACK LIBRARIES
const { useState, useEffect } = React;
const { Container, Alert } = ReactBootstrap || {};

const MainNavbar = window.MainNavbar;
const CreateProject = window.CreateProject;
const ProjectList = window.ProjectList;

function App() {
    // --- STATE ---
    const [version] = useState("React-Bootstrap v4.1.1");
    const [activeTab, setActiveTab] = useState("ProjectList");
    
    // VALIDATION STATE: Defaults to true, flips to false if check fails
    const [isValidFile, setIsValidFile] = useState(true);
    const [currentName, setCurrentName] = useState("");

    // --- INITIALIZATION & SECURITY CHECK ---
    useEffect(() => {
        Office.onReady((info) => {
            if (info.host === Office.HostType.Excel) {
                console.log("Office Ready");

                // 1. Get URL
                const url = Office.context.document.url;
                
                // 2. Decode (Fixes %20 spaces) & Clean
                const decodedUrl = decodeURI(url);
                
                // 3. Extract simple name for display (optional)
                const segments = decodedUrl.split('/');
                const fileName = segments[segments.length - 1];
                setCurrentName(fileName);

                // 4. THE CHECK
                // We check if the URL actually contains our required filename
                // If URL is empty (unsaved file), it will also fail (which is good security)
                if (!decodedUrl.includes(REQUIRED_FILENAME)) {
                    setIsValidFile(false);
                } else {
                    setIsValidFile(true);
                }
            }
        });
    }, []);

    // --- THE UI ---
    return (
        <>
            {/* Pass validation state to Navbar to hide links */}
            <MainNavbar 
                activeTab={activeTab} 
                setActiveTab={setActiveTab} 
                isFileValid={isValidFile}
            />
            
            <Container className="p-3">
                
                {/* SCENARIO A: BLOCKING ALERT */}
                {!isValidFile ? (
                    <div className="mt-4">
                        <Alert variant="danger" className="shadow-sm border-danger text-center">
                            <div className="mb-2" style={{fontSize: "1.5rem"}}>⛔</div>
                            <strong>Functionality Locked</strong>
                            <p className="small mt-2 mb-0">
                                There are no Barbizon Functions available for this Excel File.
                            </p>
                            <hr/>
                            <p className="small mb-0 text-muted">
                                Please open: <br/>
                                <span className="fw-bold text-dark">{REQUIRED_FILENAME}</span>
                            </p>
                        </Alert>
                        {/* Debug info (Optional, helpful for setup) */}
                        <div className="text-center text-muted" style={{fontSize: "0.6rem"}}>
                            Current: {currentName || "Unsaved File"}
                        </div>
                    </div>
                ) : (
                    /* VALID FILE: RENDER APP NORMALLY */
                    <>
                        {activeTab === "ProjectList" && <ProjectList />}
                        {activeTab === "AddProject" && <CreateProject />}
                    </>
                )}

                {/* Version Footer */}
                <div className="fixed-bottom p-1 text-end text-muted pe-3" style={{fontSize: "0.6rem", backgroundColor: "#f1f1f1"}}>
                    {version}
                </div>
            </Container>
        </>
    );
}

// Mount the React App
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
