/* global React, ReactDOM, Office, ReactBootstrap */

// 1. CONFIGURATION
// Change this to a unique part of your server path (e.g., "SharePoint", "OneDrive", or a specific folder name)
// If you want to disable this check for testing, leave it as an empty string ""
const EXPECTED_LOCATION = ""; 

// 2. UNPACK LIBRARIES
const { useState, useEffect } = React;
// Added 'Alert' to the unpacked list
const { Container, Alert } = ReactBootstrap || {};

const MainNavbar = window.MainNavbar;
const CreateProject = window.CreateProject;
const ProjectList = window.ProjectList;

function App() {
    // --- STATE ---
    const [version] = useState("React-Bootstrap v3.7.0");
    const [activeTab, setActiveTab] = useState("ProjectList");
    const [locationWarning, setLocationWarning] = useState(null);

    // --- INITIALIZATION & LOCATION CHECK ---
    useEffect(() => {
        Office.onReady((info) => {
            if (info.host === Office.HostType.Excel) {
                console.log("Office Ready");

                // CHECK FILE LOCATION
                const currentUrl = Office.context.document.url;
                
                // Logic: If an expected location is defined AND the current URL doesn't contain it
                if (EXPECTED_LOCATION && currentUrl && !currentUrl.includes(EXPECTED_LOCATION)) {
                    setLocationWarning({
                        title: "⚠️ Wrong File Location",
                        msg: "You are not in the Server File. Please close this and open the official Master copy."
                    });
                }
                
                // Debugging: Log the current URL so you can see what to filter for
                console.log("Current File URL:", currentUrl);
            }
        });
    }, []);

    // --- THE UI ---
    return (
        <>
            <MainNavbar activeTab={activeTab} setActiveTab={setActiveTab} />
            
            <Container className="p-3">
                
                {/* LOCATION WARNING ALERT */}
                {locationWarning && (
                    <Alert variant="warning" className="small shadow-sm border-warning">
                        <strong>{locationWarning.title}</strong><br/>
                        {locationWarning.msg}
                        <div className="mt-2 text-muted" style={{fontSize: "0.7em", wordBreak: "break-all"}}>
                            Path: {Office.context.document.url}
                        </div>
                    </Alert>
                )}

                {activeTab === "ProjectList" && <ProjectList />}
                {activeTab === "AddProject" && <CreateProject />}

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
