/* global React, ReactBootstrap, Excel */

const { useState } = React;
const { Container, Nav, Navbar, NavDropdown } = ReactBootstrap;

const MainNavbar = ({ activeTab, setActiveTab, isFileValid }) => {
    const [expanded, setExpanded] = useState(false);
    const [isSyncing, setIsSyncing] = useState(false);

    const handleLayoutChange = (newState) => {
        setExpanded(newState);
        let count = 0;
        const interval = setInterval(() => {
            window.dispatchEvent(new Event('resize'));
            count++;
            if (count > 20) clearInterval(interval);
        }, 20);
    };

    // --- NEW: TRIGGER VISUAL UPDATE ---
    const handleRefreshVisuals = async () => {
        if (isSyncing) return;
        setIsSyncing(true);
        try {
            await Excel.run(async (context) => {
                if (window.VisualLogic) {
                    await window.VisualLogic.refreshGridAlerts(context);
                }
            });
            // Simple visual feedback (alert or just logging)
            console.log("Visuals Refreshed");
        } catch (e) {
            console.error(e);
        } finally {
            setIsSyncing(false);
            setExpanded(false); // Close menu
        }
    };

    return (
        <Navbar 
            expand="lg" 
            bg="dark" 
            className="p-3 navbar-dark" 
            activeKey={activeTab} 
            expanded={expanded}
            onToggle={() => handleLayoutChange(!expanded)}
            onSelect={(selectedKey) => {
                setActiveTab(selectedKey);
                handleLayoutChange(false); 
            }}
        >
            <Navbar.Brand style={{cursor: "default"}}>
                <img 
                    src="https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg" 
                    srcSet="https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg 1x,https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg 2x,https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg 3x" 
                    alt="Barbizon Lighting Company" 
                    style={{width: "132px"}} 
                />
            </Navbar.Brand>
            
            <Navbar.Toggle aria-controls="basic-navbar-nav" />
            
            <Navbar.Collapse id="basic-navbar-nav">
                <Nav className="me-auto">
                    
                    {isFileValid ? (
                        <>
                            <Nav.Link eventKey="ProjectList">Active Projects</Nav.Link>
                            <Nav.Link eventKey="AddProject">Add a New Project</Nav.Link>
                            
                            <NavDropdown title="Options" id="basic-nav-dropdown">
                                {/* NEW BUTTON */}
                                <NavDropdown.Item onClick={handleRefreshVisuals} disabled={isSyncing}>
                                    {isSyncing ? "Syncing..." : "Refresh Grid Alerts"}
                                </NavDropdown.Item>
                                <NavDropdown.Divider />
                                <NavDropdown.Item href="#action/3.1">Export CSV</NavDropdown.Item>
                                <NavDropdown.Item href="#action/3.2">Print View</NavDropdown.Item>
                            </NavDropdown>
                        </>
                    ) : (
                        <Nav.Item className="text-muted small pt-2 fst-italic">
                           Functions Disabled
                        </Nav.Item>
                    )}

                </Nav>
            </Navbar.Collapse>
        </Navbar>
    );
}

window.MainNavbar = MainNavbar;
