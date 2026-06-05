/* global React, ReactBootstrap, Excel */
import React, { useState } from 'react';
import { Container, Nav, Navbar, NavDropdown, Badge, Modal, Form, Button, Spinner } from 'react-bootstrap';

import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faBell, faPaintRoller, faUserCog, faFileExport, faPrint, faBars, faUserCircle, faCodeCompare, faTimes, faTerminal, faCalendarAlt, faPlus } from '@fortawesome/free-solid-svg-icons';
import { VisualLogic } from '../utils/visualLogic';
import { FormattingLogic } from '../utils/formattingLogic_v2';
import { ChangelogLogic } from '../utils/changelogLogic';
import { IdentityLogic } from '../utils/identityLogic';
import { faCircleUser } from '@fortawesome/free-regular-svg-icons';

const AUTHORIZED_ADMINS = ["Rob", "Kevin", "Rob Kreps", "Kevin Rittner", "Jason", "Jason Mace"];

const MainNavbar = ({ activeTab, setActiveTab, isFileValid, unseenCount, fileType, onAddTimecard }) => {
    const [expanded, setExpanded] = useState(false);
    const [isSyncing, setIsSyncing] = useState(false);

    const [showVersionModal, setShowVersionModal] = useState(false);
    const [versionInput, setVersionInput] = useState("");

    const triggerResizeLoop = () => {
        let count = 0;
        const interval = setInterval(() => {
            window.dispatchEvent(new Event('resize'));
            count++;
            if (count > 25) clearInterval(interval);
        }, 20);
    };

    const handleLayoutChange = (newState) => {
        setExpanded(newState);
        triggerResizeLoop();
    };

    // --- 1. TRIGGER GRID ALERTS (VisualLogic) ---
    const handleRefreshVisuals = async () => {
        if (isSyncing) return;
        handleLayoutChange(false); // Collapse menu immediately
        
        // Small delay to allow the menu animation to start before heavy logic
        await new Promise(resolve => setTimeout(resolve, 150));

        setIsSyncing(true);
        try {
            await Excel.run(async (context) => {
                for (const name of ["Houston", "Dallas"]) {
                    await VisualLogic.refreshGridAlerts(context, name);
                }
            });
            console.log("Visuals Refreshed");
        } catch (e) {
            console.error(e);
        } finally {
            setIsSyncing(false);
        }
    };

    // --- 2. TRIGGER FORMATTING RESET (FormattingLogic) ---
    const handleResetFormatting = async () => {
        if (isSyncing) return;
        handleLayoutChange(false); // Collapse menu immediately

        // Small delay to allow the menu animation to start before heavy logic
        await new Promise(resolve => setTimeout(resolve, 150));

        setIsSyncing(true);
        try {
            await Excel.run(async (context) => {
                for (const name of ["Houston", "Dallas"]) {
                    await FormattingLogic.generateSmartRules(context, name);
                }
            });
            console.log("Formatting Reset");
        } catch (e) {
            console.error(e);
        } finally {
            setIsSyncing(false);
        }
    };

    // --- 3. TOGGLE CHANGELOG VISIBILITY (Developer Only) ---
    const handleToggleChangelog = async () => {
        if (isSyncing) return;
        handleLayoutChange(false);

        setIsSyncing(true);
        try {
            await Excel.run(async (context) => {
                const sheetNames = ["changelog", "Changelog", "Change Log"];
                let targetSheet = null;

                for (const name of sheetNames) {
                    const sheet = context.workbook.worksheets.getItemOrNullObject(name);
                    sheet.load("visibility, name");
                    await context.sync();
                    if (!sheet.isNullObject) {
                        targetSheet = sheet;
                        break;
                    }
                }

                if (targetSheet) {
                    const isVisible = targetSheet.visibility === Excel.SheetVisibility.visible;
                    targetSheet.visibility = isVisible ? Excel.SheetVisibility.veryHidden : Excel.SheetVisibility.visible;
                    await context.sync();
                    if (window.GlobalToast) window.GlobalToast.success(`Sheet is now ${!isVisible ? "Visible" : "Very Hidden"}`);
                } else {
                    if (window.GlobalToast) window.GlobalToast.error("Changelog sheet not found.");
                }
            });
        } catch (e) {
            console.error(e);
        } finally {
            setIsSyncing(false);
        }
    };

    // --- 4. TOGGLE VACATIONS VISIBILITY (Developer Only) ---
    const handleToggleVacations = async () => {
        if (isSyncing) return;
        handleLayoutChange(false);

        setIsSyncing(true);
        try {
            await Excel.run(async (context) => {
                const sheetNames = ["Vacations", "Vacation-PTO", "PTO"];
                let targetSheet = null;

                for (const name of sheetNames) {
                    const sheet = context.workbook.worksheets.getItemOrNullObject(name);
                    sheet.load("visibility, name");
                    await context.sync();
                    if (!sheet.isNullObject) {
                        targetSheet = sheet;
                        break;
                    }
                }

                if (targetSheet) {
                    const isVisible = targetSheet.visibility === Excel.SheetVisibility.visible;
                    targetSheet.visibility = isVisible ? Excel.SheetVisibility.veryHidden : Excel.SheetVisibility.visible;
                    await context.sync();
                    if (window.GlobalToast) window.GlobalToast.success(`${targetSheet.name} is now ${!isVisible ? "Visible" : "Very Hidden"}`);
                } else {
                    if (window.GlobalToast) window.GlobalToast.error("Vacations sheet not found.");
                }
            });
        } catch (e) {
            console.error(e);
        } finally {
            setIsSyncing(false);
        }
    };

    // --- 5. TOGGLE TEAM VISIBILITY (Developer Only) ---
    const handleToggleTeam = async () => {
        if (isSyncing) return;
        handleLayoutChange(false);

        setIsSyncing(true);
        try {
            await Excel.run(async (context) => {
                const sheetNames = ["Team", "TEAM"];
                let targetSheet = null;

                for (const name of sheetNames) {
                    const sheet = context.workbook.worksheets.getItemOrNullObject(name);
                    sheet.load("visibility, name");
                    await context.sync();
                    if (!sheet.isNullObject) {
                        targetSheet = sheet;
                        break;
                    }
                }

                if (targetSheet) {
                    const isVisible = targetSheet.visibility === Excel.SheetVisibility.visible;
                    targetSheet.visibility = isVisible ? Excel.SheetVisibility.veryHidden : Excel.SheetVisibility.visible;
                    await context.sync();
                    if (window.GlobalToast) window.GlobalToast.success(`${targetSheet.name} is now ${!isVisible ? "Visible" : "Very Hidden"}`);
                } else {
                    if (window.GlobalToast) window.GlobalToast.error("Team sheet not found.");
                }
            });
        } catch (e) {
            console.error(e);
        } finally {
            setIsSyncing(false);
        }
    };

    const handleLogVersion = async () => {
        if (!versionInput) return;
        handleLayoutChange(false);

        setIsSyncing(true);
        try {
            await Excel.run(async (context) => {
                // We log with "Admin" override so it triggers the global toast for all users
                await ChangelogLogic.logChange(context, versionInput, "Admin");
            });
            if (window.GlobalToast) window.GlobalToast.success("Version entry logged as Admin.");
            setVersionInput("");
            setShowVersionModal(false);
        } catch (e) {
            console.error(e);
        } finally {
            setIsSyncing(false);
        }
    };

    return (
        <>
        <Navbar 
            expand="lg" 
            bg="dark" 
            className="p-3 navbar-dark flex-shrink-0"
            activeKey={activeTab} 
            expanded={expanded} 
            onToggle={() => handleLayoutChange(!expanded)}
            onSelect={(selectedKey) => {
                // Sync Excel worksheet activation with menu selection
                if (selectedKey === "HoustonList" || selectedKey === "DallasList") {
                    const targetSheet = selectedKey === "HoustonList" ? "Houston" : "Dallas";
                    Excel.run(async (context) => {
                        const sheet = context.workbook.worksheets.getItem(targetSheet);
                        sheet.activate();
                        await context.sync();
                    }).catch(err => console.warn(`Navigation: Could not activate sheet "${targetSheet}".`, err));
                }

                // Only switch tabs if it's a valid eventKey (not undefined)
                if (selectedKey) setActiveTab(selectedKey);
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

            {/* USER ICON & TOGGLE GROUP (Always on the right) */}
            <div className="d-flex align-items-center order-lg-last">
                {isFileValid && (
                    <NavDropdown 
                        align="end"
                        title={
                            <div className="position-relative d-inline-block">
                                <FontAwesomeIcon icon={faCircleUser} size="lg" className="text-white opacity-75" />
                                {unseenCount > 0 && (
                                    <Badge 
                                        pill bg="danger" 
                                        className="position-absolute top-0 start-75 translate-middle"
                                        style={{ fontSize: '0.6rem', padding: '0.25em 0.5em' }}
                                    >
                                        {unseenCount}
                                    </Badge>
                                )}
                            </div>
                        } 
                        id="user-nav-dropdown"
                        className="me-2 no-caret"
                        onToggle={() => triggerResizeLoop()}
                    >
                        <NavDropdown.Item eventKey="Updates" className="d-flex justify-content-between align-items-center">
                            <span><FontAwesomeIcon icon={faCodeCompare} className="me-2 text-muted" />Change History</span>
                            {unseenCount > 0 && <Badge bg="danger" pill className="ms-2">{unseenCount}</Badge>}
                        </NavDropdown.Item>
                        <NavDropdown.Item eventKey="Settings">
                            <FontAwesomeIcon icon={faUserCog} className="me-2 text-muted" /> User Settings
                        </NavDropdown.Item>
                    </NavDropdown>
                )}
                
                <Navbar.Toggle aria-controls="basic-navbar-nav" />
            </div>

            <Navbar.Collapse id="basic-navbar-nav">
                <Nav className="me-auto">
                    {isFileValid ? (
                        <>
                            {fileType === 'gantt' && (
                                <>
                                    <Nav.Link eventKey="Home">Home</Nav.Link>
                                    <Nav.Link eventKey="HoustonList">Houston Projects</Nav.Link>
                                    <Nav.Link eventKey="DallasList">Dallas Projects</Nav.Link>
                                    <Nav.Link eventKey="AddProject">Add a New Project</Nav.Link>
                                    <Nav.Link eventKey="SubContractorManager">Sub Contractors</Nav.Link>

                                    <NavDropdown title="PTO/Vacation" id="pto-nav-dropdown">
                                        <NavDropdown.Item eventKey="PTOManager">Manage PTO List</NavDropdown.Item>
                                        <NavDropdown.Item eventKey="AddPTO">Add PTO / Vacation</NavDropdown.Item>
                                    </NavDropdown>
                                    
                                    <NavDropdown 
                                        title="Options" 
                                        id="basic-nav-dropdown" 
                                        onToggle={() => triggerResizeLoop()}
                                    >
                                        {/* ACTION 1: GRID ALERTS */}
                                        <NavDropdown.Item onClick={handleRefreshVisuals} disabled={isSyncing} aria-label="Refresh Grid Alerts">
                                            <FontAwesomeIcon icon={faBell} className="me-2 text-muted" /> 
                                            {isSyncing ? "Syncing..." : "Refresh Grid Alerts"}
                                        </NavDropdown.Item>

                                        {/* ACTION 2: FORMATTING */}
                                        <NavDropdown.Item onClick={handleResetFormatting} disabled={isSyncing} aria-label="Reset Formatting Rules">
                                            <FontAwesomeIcon icon={faPaintRoller} className="me-2 text-muted" /> 
                                            Reset Formatting Rules
                                        </NavDropdown.Item>
                                        
                                        {AUTHORIZED_ADMINS.includes(IdentityLogic.getIdentity()) && (
                                            <NavDropdown.Item eventKey="TeamManager">
                                                <FontAwesomeIcon icon={faUserCog} className="me-2 text-muted" /> 
                                                Team Management
                                            </NavDropdown.Item>
                                        )}

                                        <NavDropdown.Divider />
                                        
                                        <NavDropdown.Item href="#action/3.1" className="d-flex justify-content-between align-items-center" disabled>
                                            <span><FontAwesomeIcon icon={faFileExport} className="me-2 text-muted" /> Export CSV</span>
                                            <Badge bg="secondary" pill className="ms-2" style={{ fontSize: '0.6rem' }}>Coming Soon</Badge>
                                        </NavDropdown.Item>
                                        <NavDropdown.Item href="#action/3.2" className="d-flex justify-content-between align-items-center" disabled>
                                            <span><FontAwesomeIcon icon={faPrint} className="me-2 text-muted" /> Print View</span>
                                            <Badge bg="secondary" pill className="ms-2" style={{ fontSize: '0.6rem' }}>Coming Soon</Badge>
                                        </NavDropdown.Item>
                                    </NavDropdown>

                                    {/* DEVELOPER MENU (Jason Only) */}
                                    {IdentityLogic.getIdentity() === "Jason" && (
                                        <NavDropdown title="Developer" id="dev-nav-dropdown" onToggle={() => triggerResizeLoop()}>
                                            <NavDropdown.Item onClick={handleToggleChangelog}>
                                                <FontAwesomeIcon icon={faTerminal} className="me-2 text-muted" /> Toggle Changelog
                                            </NavDropdown.Item>
                                            <NavDropdown.Item onClick={handleToggleVacations}>
                                                <FontAwesomeIcon icon={faCalendarAlt} className="me-2 text-muted" /> Toggle Vacations
                                            </NavDropdown.Item>
                                            <NavDropdown.Item onClick={handleToggleTeam}>
                                                <FontAwesomeIcon icon={faUserCog} className="me-2 text-muted" /> Toggle Team
                                            </NavDropdown.Item>
                                            <NavDropdown.Divider />
                                            <NavDropdown.Item onClick={() => { setShowVersionModal(true); setExpanded(false); }}>
                                                <FontAwesomeIcon icon={faPlus} className="me-2 text-muted" /> Log Version Entry
                                            </NavDropdown.Item>
                                        </NavDropdown>
                                    )}
                                </>
                            )}

                            {fileType === 'timecard' && (
                                <>
                                    <Nav.Link eventKey="TimecardDashboard">Timecard Dashboard</Nav.Link>
                                    <Nav.Link onClick={() => onAddTimecard()}>Add Timesheet</Nav.Link>
                                </>
                            )}

                            {fileType === 'pm_timelog' && (
                                <>
                                    <Nav.Link eventKey="PmTimelogDashboard">Timelog Dashboard</Nav.Link>
                                </>
                            )}
                        </>
                    ) : (
                        <Nav.Item className="text-danger small pt-2 fw-bold">
                            <FontAwesomeIcon icon={faTimes} className="me-2" /> Helpers Inactive
                        </Nav.Item>
                    )}
                </Nav>
            </Navbar.Collapse>
        </Navbar>

        {/* VERSION ENTRY MODAL */}
        <Modal show={showVersionModal} onHide={() => setShowVersionModal(false)} centered>
            <Modal.Header closeButton className="py-2 bg-light">
                <Modal.Title style={{fontSize: "1rem"}} className="fw-bold text-primary">Log Version/Admin Update</Modal.Title>
            </Modal.Header>
            <Modal.Body className="p-3">
                <Form.Group>
                    <Form.Label className="small fw-bold text-muted text-uppercase">Update Description</Form.Label>
                    <Form.Control 
                        as="textarea" 
                        rows={3} 
                        placeholder="e.g. Version 6.1: Added new Team Management features..." 
                        value={versionInput}
                        onChange={(e) => setVersionInput(e.target.value)}
                    />
                    <Form.Text className="text-muted small">This will trigger a blue notification for all users.</Form.Text>
                </Form.Group>
            </Modal.Body>
            <Modal.Footer className="py-1 bg-light">
                <Button variant="secondary" size="sm" onClick={() => setShowVersionModal(false)}>Cancel</Button>
                <Button variant="primary" size="sm" onClick={handleLogVersion} disabled={isSyncing || !versionInput}>
                    {isSyncing ? <Spinner animation="border" size="sm" /> : "Log Update"}
                </Button>
            </Modal.Footer>
        </Modal>
        </>
    );
}

export default MainNavbar;
