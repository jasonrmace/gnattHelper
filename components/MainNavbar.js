/* global React, ReactBootstrap */

const { useState } = React;
const { Container, Nav, Navbar, NavDropdown } = ReactBootstrap;

// ADDED PROP: isFileValid
const MainNavbar = ({ activeTab, setActiveTab, isFileValid }) => {
    const [expanded, setExpanded] = useState(false);

    const handleLayoutChange = (newState) => {
        setExpanded(newState);
        
        let count = 0;
        const interval = setInterval(() => {
            window.dispatchEvent(new Event('resize'));
            count++;
            if (count > 20) clearInterval(interval);
        }, 20);
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
            {/* <Container> */}
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
                    
                    {/* CONDITIONAL RENDERING BASED ON FILE NAME */}
                    {isFileValid ? (
                        <>
                            <Nav.Link eventKey="ProjectList">Active Projects</Nav.Link>
                            <Nav.Link eventKey="AddProject">Add a New Project</Nav.Link>
                            
                            <NavDropdown title="Options" id="basic-nav-dropdown">
                                <NavDropdown.Item href="#action/3.1">Export CSV</NavDropdown.Item>
                                <NavDropdown.Item href="#action/3.2">Print View</NavDropdown.Item>
                                <NavDropdown.Divider />
                                <NavDropdown.Item href="#action/3.4">Settings</NavDropdown.Item>
                            </NavDropdown>
                        </>
                    ) : (
                        <Nav.Item className="text-muted small pt-2 fst-italic">
                           Functions Disabled
                        </Nav.Item>
                    )}

                </Nav>
            </Navbar.Collapse>
            {/* </Container> */}
        </Navbar>
    );
}

window.MainNavbar = MainNavbar;
