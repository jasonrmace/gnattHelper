/* global React, ReactDOM, Office, Excel */

import CreateProject from "./components/CreateProject";

const { useState, useEffect } = React;

function App() {
    // --- STATE MANAGEMENT ---
    const [version] = useState("React Split v2.0"); // Verify this in the UI

    // --- INITIALIZATION ---
    useEffect(() => {
        Office.onReady((info) => {
            if (info.host === Office.HostType.Excel) {
                console.log("Office Ready - React Loaded");
            }
        });
    }, []);

    // --- THE UI RENDER ---
    return (
        <div className="container-fluid p-3">
            
            <div className="border-bottom border-primary pb-2 mb-3">
                <h5 className="text-primary fw-bold m-0">Gantt Manager</h5>
                <small className="text-muted" style={{fontSize: "0.7rem"}}>React Edition</small>
            </div>

            <CreateProject />
            
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
