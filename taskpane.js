// 1. UNPACK LIBRARIES (The "No-Import" Way)
const { useState, useEffect } = React;
const { Container } = ReactBootstrap;

const MainNavbar = window.MainNavbar;
const CreateProject = window.CreateProject;

function App() {
    // --- STATE ---
    const [version] = useState("React-Bootstrap v3.1");

    // --- INITIALIZATION ---
    useEffect(() => {
        Office.onReady((info) => {
            if (info.host === Office.HostType.Excel) {
                console.log("Office Ready");
            }
        });
    }, []);

    // --- THE UI (Using React-Bootstrap Components) ---
    return (
        <Container className="p-3">
            
            <MainNavbar />
            
            <div className="border-bottom border-primary pb-2 mb-3">
                <h5 className="text-primary fw-bold m-0">Gantt Manager</h5>
                <small className="text-muted" style={{fontSize: "0.7rem"}}>RB Edition</small>
            </div>

            <CreateProject />
            
            {/* Version Footer */}
             <div className="fixed-bottom p-1 text-end text-muted pe-3" style={{fontSize: "0.6rem", backgroundColor: "#f1f1f1"}}>
                {version}
            </div>

        </Container>
    );
}

// Mount the React App
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
