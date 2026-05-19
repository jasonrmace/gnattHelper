const { Container, Nav, Navbar, NavDropdown } = ReactBootstrap;

const MainNavbar = () => {
    return (
    <Navbar expand="lg" bg="dark">
      {/* <Container> */}
        <Navbar.Brand href="#home">
            <img src="https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg" srcSet="https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg 1x,https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg 2x,https://barbizon.com/wp-content/uploads/2022/12/barbizon-logo-22.svg 3x" alt="Barbizon Lighting Company" style={{width: "132px"}} />
        </Navbar.Brand>
        <Navbar.Toggle aria-controls="basic-navbar-nav" />
        <Navbar.Collapse id="basic-navbar-nav">
          <Nav className="me-auto">
            <Nav.Link href="#home">Home</Nav.Link>
            <Nav.Link href="#link">Link</Nav.Link>
            <NavDropdown title="Dropdown" id="basic-nav-dropdown">
              <NavDropdown.Item href="#action/3.1">Action</NavDropdown.Item>
              <NavDropdown.Item href="#action/3.2">
                Another action
              </NavDropdown.Item>
              <NavDropdown.Item href="#action/3.3">Something</NavDropdown.Item>
              <NavDropdown.Divider />
              <NavDropdown.Item href="#action/3.4">
                Separated link
              </NavDropdown.Item>
            </NavDropdown>
          </Nav>
        </Navbar.Collapse>
      {/* </Container> */}
    </Navbar>
  );
}

window.MainNavbar = MainNavbar;