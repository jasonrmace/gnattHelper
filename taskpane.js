// 1. UNPACK LIBRARIES (The "No-Import" Way)
const { useState, useEffect } = React;
const { Container } = ReactBootstrap;

const MainNavbar = window.MainNavbar;
const CreateProject = window.CreateProject;
const ProjectList = window.ProjectList;

function App() {
    // --- STATE ---
    const [version] = useState("React-Bootstrap v3.1");
    const [activeTab, setActiveTab] = useState("home");

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
        <>
            <MainNavbar activeTab={activeTab} setActiveTab={setActiveTab} />
            <Container className="p-3">

                {activeTab === "home" && <ProjectList />}
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
