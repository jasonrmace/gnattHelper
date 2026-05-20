/* global React, ReactBootstrap */

const { useState } = React;
const { Container, Nav, Navbar, NavDropdown } = ReactBootstrap;

const MainNavbar = ({ activeTab, setActiveTab }) => {
    // 1. State to control the collapse
    const [expanded, setExpanded] = useState(false);

    return (
        <Navbar 
            expand="lg" 
            bg="dark" 
            className="p-3 navbar-dark" 
            activeKey={activeTab} 
            
            // 2. Bind the state to the component
            expanded={expanded}
            
            // 3. Handle the "Hamburger" click
            onToggle={() => setExpanded(!expanded)}
            
            // 4. When a link is clicked, switch tabs AND close menu
            onSelect={(selectedKey) => {
                setActiveTab(selectedKey);
                setExpanded(false); 
            }}
        >
            {/* <Container> */}
            <Navbar.Brand eventKey="ProjectList" onClick={() => setActiveTab("ProjectList")}>
                <img 
                    src="https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg" 
                    srcSet="https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg 1x,https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg 2x,https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg 3x" 
                    alt="Barbizon Lighting Company" 
                    style={{width: "132px", cursor: "pointer"}} 
                />
            </Navbar.Brand>
            
            <Navbar.Toggle aria-controls="basic-navbar-nav" />
            
            <Navbar.Collapse id="basic-navbar-nav">
                <Nav className="me-auto">
                    <Nav.Link eventKey="ProjectList">Active Projects</Nav.Link>
                    <Nav.Link eventKey="AddProject">Add a New Project</Nav.Link>
                    
                    <NavDropdown title="Dropdown" id="basic-nav-dropdown">
                        <NavDropdown.Item eventKey="action-1" href="#action/3.1">Action</NavDropdown.Item>
                        <NavDropdown.Item eventKey="action-2" href="#action/3.2">Another action</NavDropdown.Item>
                        <NavDropdown.Item eventKey="action-3" href="#action/3.3">Something</NavDropdown.Item>
                        <NavDropdown.Divider />
                        <NavDropdown.Item eventKey="action-4" href="#action/3.4">Separated link</NavDropdown.Item>
                    </NavDropdown>
                </Nav>
            </Navbar.Collapse>
            {/* </Container> */}
        </Navbar>
    );
}

window.MainNavbar = MainNavbar;
