/* global React, ReactBootstrap, Excel */
import React, { useState } from 'react';
import { Container, Nav, Navbar, NavDropdown, Badge } from 'react-bootstrap';

import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faBell, faPaintRoller, faUserCog, faFileExport, faPrint, faBars, faUserCircle, faCodeCompare, faTimes } from '@fortawesome/free-solid-svg-icons';
import { VisualLogic } from '../utils/visualLogic';
import { FormattingLogic } from '../utils/formattingLogic_v2';
import { faCircleUser } from '@fortawesome/free-regular-svg-icons';

const MainNavbar = ({ activeTab, setActiveTab, isFileValid, unseenCount, fileType, onAddTimecard }) => {
    const [expanded, setExpanded] = useState(false);
    const [isSyncing, setIsSyncing] = useState(false);

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
        setIsSyncing(true);
        try {
            await Excel.run(async (context) => {
                await VisualLogic.refreshGridAlerts(context);
            });
            console.log("Visuals Refreshed");
        } catch (e) {
            console.error(e);
        } finally {
            setIsSyncing(false);
            setExpanded(false);
        }
    };

    // --- 2. TRIGGER FORMATTING RESET (FormattingLogic) ---
    const handleResetFormatting = async () => {
        if (isSyncing) return;
        setIsSyncing(true);
        try {
            await Excel.run(async (context) => {
                await FormattingLogic.generateSmartRules(context);
            });
            console.log("Formatting Reset");
        } catch (e) {
            console.error(e);
        } finally {
            setIsSyncing(false);
            setExpanded(false);
        }
    };

    return (
        <Navbar 
            expand="lg" 
            bg="dark" 
            className="p-3 navbar-dark flex-shrink-0"
            activeKey={activeTab} 
            expanded={expanded} 
            onToggle={() => handleLayoutChange(!expanded)}
            onSelect={(selectedKey) => {
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
                                    <Nav.Link eventKey="ProjectList">Active Projects</Nav.Link>
                                    <Nav.Link eventKey="AddProject">Add a New Project</Nav.Link>
                                    
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
                                        
                                        <NavDropdown.Divider />
                                        
                                        <NavDropdown.Item href="#action/3.1"><FontAwesomeIcon icon={faFileExport} className="me-2 text-muted" /> Export CSV</NavDropdown.Item>
                                        <NavDropdown.Item href="#action/3.2"><FontAwesomeIcon icon={faPrint} className="me-2 text-muted" /> Print View</NavDropdown.Item>
                                    </NavDropdown>
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
    );
}

export default MainNavbar;
