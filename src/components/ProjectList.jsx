/* global React, ReactBootstrap, Excel */
import ProjectTasks from './ProjectTasks';

import React, { useState, useEffect, useRef } from 'react';
import { Button, Card, Badge, Spinner, Row, Col, Form, Collapse } from 'react-bootstrap';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faLocationArrow, faListCheck, faChevronRight, faUser, faCalendarDays, faArrowRight, faSyncAlt, faSortAmountDown, faSortAmountUp, faSliders } from '@fortawesome/free-solid-svg-icons';

const ProjectList = ({ sheetName = "Houston", refreshTrigger, highlightId }) => {
    // --- STATE ---
    const [projects, setProjects] = useState([]);
    const [teamMembers, setTeamMembers] = useState([]);
    const [isFetching, setIsFetching] = useState(false);
    
    const [filters, setFilters] = useState({ lead: "", percent: "" });
    const [sortConfig, setSortConfig] = useState({ key: "id", direction: "asc" });
    const [showControls, setShowControls] = useState(false);
    
    // VIEW STATE: Controls "List" vs "Detail" view
    const [selectedProject, setSelectedProject] = useState(null);

    // Scrolling Ref
    const projectRefs = useRef({});

    // Dynamic Height State
    const listContainerRef = useRef(null);
    const [listHeight, setListHeight] = useState("500px");

    // --- 1. DYNAMIC HEIGHT CALCULATION ---
    // Runs on load AND when switching views
    useEffect(() => {
        const calculateHeight = () => {
            if (listContainerRef.current) {
                const topPosition = listContainerRef.current.getBoundingClientRect().top;
                const availableHeight = window.innerHeight - topPosition - 20 - 34;
                setListHeight(`${availableHeight}px`);
            }
        };
        calculateHeight();
        window.addEventListener('resize', calculateHeight);
        return () => window.removeEventListener('resize', calculateHeight);
    }, [selectedProject]); 

    // --- 2. SCROLL TO HIGHLIGHTED PROJECT ---
    useEffect(() => {
        if (highlightId && projectRefs.current[highlightId]) {
            // Small timeout to ensure the list has finished rendering after a refresh
            setTimeout(() => {
                projectRefs.current[highlightId].scrollIntoView({ behavior: "smooth", block: "center" });
            }, 500);
        }
    }, [highlightId, projects]);

    // --- 2. EXCEL ACTIONS ---
    const handleJump = async (rowIndex) => {
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(sheetName);
                sheet.activate();
                const range = sheet.getRangeByIndexes(rowIndex, 0, 1, 1).getEntireRow();
                range.select();
                await context.sync();
            });
        } catch (error) { console.error("Jump Error:", error); }
    };

    // --- 3. DATA FETCHING ---
    const fetchProjects = async () => {
        setIsFetching(true);
        try {
            await Excel.run(async (context) => {
                const teamSheet = context.workbook.worksheets.getItemOrNullObject("Team");
                const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                
                teamSheet.load("isNullObject");
                sheet.load("isNullObject");
                await context.sync();

                // Handle cases where sheets are missing (e.g. initial load in a Timecard file)
                if (teamSheet.isNullObject || sheet.isNullObject) {
                    setProjects([]);
                    return;
                }

                const teamRange = teamSheet.getUsedRange();
                teamRange.load("text");

                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                footerRange.load(["rowIndex", "isNullObject"]);
                await context.sync();

                const teamMap = {};
                const teamRows = teamRange.text;
                const members = [];
                for (let i = 1; i < teamRows.length; i++) {
                    const firstName = teamRows[i]?.[0]?.trim() || "";
                    const lastName = teamRows[i]?.[1]?.trim() || "";
                    if (firstName) {
                        const full = `${firstName} ${lastName}`.trim();
                        teamMap[firstName.toLowerCase()] = full;
                        members.push({ first: firstName, full: full });
                    }
                }
                setTeamMembers(members);

                const dataStartIndex = 7; 
                if (footerRange.isNullObject) {
                    setProjects([]);
                    return;
                }

                const footerIndex = footerRange.rowIndex;
                const rowCount = footerIndex - dataStartIndex;

                if (rowCount <= 0) {
                    setProjects([]); 
                    return;
                }

                // Find Project Number Column (Dynamic Lookup in Row 7)
                const headerRow = sheet.getRange("7:7");
                const projNumHeader = headerRow.find("Project Number", { completeMatch: true, matchCase: false });
                projNumHeader.load(["columnIndex", "isNullObject"]);
                await context.sync();

                let projNumIdx = 3; // Fallback to index 3 (Col D)
                if (!projNumHeader.isNullObject) {
                    projNumIdx = projNumHeader.columnIndex;
                }

                const colCountToFetch = Math.max(8, projNumIdx + 1);
                const dataRange = sheet.getRangeByIndexes(dataStartIndex, 0, rowCount, colCountToFetch);
                dataRange.load("text"); 
                await context.sync();

                const rawRows = dataRange.text;
                const projectsMap = new Map();

                // Aggregate Projects & Tasks
                rawRows.forEach((row, index) => {
                    if (!row[1] || row[1] === "") return;
                    const id = parseFloat(row[0]);
                    const currentRowIndex = dataStartIndex + index; 

                    if (!isNaN(id) && Number.isInteger(id)) {
                        // IT IS A PROJECT
                        const rawLead = row[2]?.trim() || "";
                        const fullLeadName = teamMap[rawLead.toLowerCase()] || rawLead;
                        projectsMap.set(id, {
                            location: sheetName,
                            id: row[0], // Keep as string/raw to avoid float issues
                            projectNumber: row[projNumIdx],
                            rowIndex: currentRowIndex,
                            name: row[1],
                            lead: fullLeadName,
                            start: row[4],
                            end: row[5],
                            percent: row[7],
                            totalTasks: 0,
                            completedTasks: 0
                        });
                    } else if (!isNaN(id) && !Number.isInteger(id)) {
                        // IT IS A TASK
                        const parentId = Math.floor(id);
                        if (projectsMap.has(parentId)) {
                            const project = projectsMap.get(parentId);
                            project.totalTasks++;
                            if (row[7].includes("100%")) project.completedTasks++;
                        }
                    }
                });
                setProjects(Array.from(projectsMap.values()));
            });
        } catch (error) { console.error(error); } finally { setIsFetching(false); }
    };

    useEffect(() => { fetchProjects(); }, [refreshTrigger]);

    // --- 4. FILTER & SORT LOGIC ---
    const getProcessedProjects = () => {
        let filtered = projects.filter(p => {
            const matchLead = !filters.lead || p.lead === filters.lead;
            const matchPercent = !filters.percent || p.percent === filters.percent;
            return matchLead && matchPercent;
        });

        return filtered.sort((a, b) => {
            let valA = a[sortConfig.key];
            let valB = b[sortConfig.key];

            if (sortConfig.key === 'id') {
                valA = parseFloat(valA);
                valB = parseFloat(valB);
            } else if (sortConfig.key === 'percent') {
                valA = parseInt(valA?.replace('%', '') || '0');
                valB = parseInt(valB?.replace('%', '') || '0');
            } else if (sortConfig.key === 'start' || sortConfig.key === 'end') {
                valA = new Date(valA || 0).getTime();
                valB = new Date(valB || 0).getTime();
            } else {
                valA = (valA || "").toString().toLowerCase();
                valB = (valB || "").toString().toLowerCase();
            }

            if (valA < valB) return sortConfig.direction === 'asc' ? -1 : 1;
            if (valA > valB) return sortConfig.direction === 'asc' ? 1 : -1;
            return 0;
        });
    };

    const processedProjects = getProcessedProjects();

    // --- VIEW ROUTER (The Switch) ---
    // If a project is selected, SHOW TASKS VIEW instead of List
    if (selectedProject) {
        return (
            <ProjectTasks 
                project={selectedProject} 
                onBack={() => setSelectedProject(null)} 
            />
        );
    }

    // --- DEFAULT UI: PROJECT LIST ---
    return (
        <div className="mt-4">
            <div className="d-flex justify-content-between align-items-center mb-2">
                <h6 className="m-0 fw-bold text-primary">{sheetName} Active Projects ({projects.length})</h6>
                <div className="d-flex align-items-center">
                    <Button variant="link" size="sm" className="text-decoration-none p-0 me-3 text-muted" onClick={() => setShowControls(!showControls)} title="Sort & Filter">
                        <FontAwesomeIcon icon={faSliders} className={showControls ? "text-primary" : ""} />
                    </Button>
                    <Button variant="link" size="sm" className="text-decoration-none p-0" onClick={fetchProjects}>
                        {isFetching ? <Spinner animation="border" size="sm" /> : <><FontAwesomeIcon icon={faSyncAlt} className="me-1" /> Refresh</>}
                    </Button>
                </div>
            </div>

            {/* CONTROLS: FILTER & SORT */}
            <Collapse in={showControls}>
                <div>
                    <div className="bg-light p-2 rounded mb-3 border shadow-sm small">
                        <Row className="g-2">
                            <Col xs={4}>
                                <Form.Label className="small fw-bold text-muted mb-1 text-uppercase" style={{ fontSize: '0.65rem' }}>Filter Lead</Form.Label>
                                <Form.Select size="sm" value={filters.lead} onChange={(e) => setFilters({...filters, lead: e.target.value})}>
                                    <option value="">All Leads</option>
                                    {teamMembers.map((m, i) => <option key={i} value={m.full}>{m.full}</option>)}
                                </Form.Select>
                            </Col>
                            <Col xs={4}>
                                <Form.Label className="small fw-bold text-muted mb-1 text-uppercase" style={{ fontSize: '0.65rem' }}>Filter Status</Form.Label>
                                <Form.Select size="sm" value={filters.percent} onChange={(e) => setFilters({...filters, percent: e.target.value})}>
                                    <option value="">All %</option>
                                    {["0%", "25%", "50%", "75%", "100%"].map(v => <option key={v} value={v}>{v}</option>)}
                                </Form.Select>
                            </Col>
                            <Col xs={4}>
                                <Form.Label className="small fw-bold text-muted mb-1 text-uppercase" style={{ fontSize: '0.65rem' }}>Sort By</Form.Label>
                                <div className="d-flex">
                                <Form.Select 
                                    size="sm" 
                                    className="me-1"
                                    value={sortConfig.key} 
                                    onChange={(e) => setSortConfig({...sortConfig, key: e.target.value})}
                                >
                                    <option value="id">ID</option>
                                    <option value="projectNumber">Proj #</option>
                                    <option value="name">Name</option>
                                    <option value="start">Start</option>
                                    <option value="end">End</option>
                                    <option value="percent">Complete</option>
                                </Form.Select>
                                <Button 
                                    variant="outline-secondary" 
                                    size="sm" 
                                    onClick={() => setSortConfig({...sortConfig, direction: sortConfig.direction === 'asc' ? 'desc' : 'asc'})}
                                >
                                    <FontAwesomeIcon icon={sortConfig.direction === 'asc' ? faSortAmountUp : faSortAmountDown} />
                                </Button>
                                </div>
                            </Col>
                        </Row>
                    </div>
                </div>
            </Collapse>

            {processedProjects.length === 0 && !isFetching && (
                <div className="text-center text-muted small mt-4">
                    No projects found matching these criteria.<br/>
                    <Button variant="link" size="sm" className="p-0 mt-1 text-decoration-none fw-bold" onClick={() => setFilters({ lead: "", percent: "" })}>
                        Clear all filters
                    </Button>
                </div>
            )}

            <div 
                ref={listContainerRef}
                style={{ maxHeight: listHeight, overflowY: "auto", transition: "max-height 0.1s ease-out" }}
            >
                {processedProjects.map((p, index) => {
                    const isNew = highlightId && String(p.id) === String(highlightId);
                    return (
                        <Card 
                            key={index} 
                            ref={el => (projectRefs.current[p.id] = el)}
                            className={`mb-2 shadow-sm border-0 ${isNew ? 'border border-primary' : ''}`}
                            style={isNew ? { backgroundColor: '#f0f7ff' } : {}}
                        >
                            <Card.Body className="py-2">
                                <Row>
                                    <Col xs="auto" className="">
                                            <Badge bg="primary">#{p.id}</Badge>
                                    </Col>
                                    <Col className="fw-bold text-dark text-truncate p-0 d-flex align-items-center" title={p.name}>
                                        <span className="text-truncate">{p.name}</span>
                                        {isNew && <Badge bg="success" className="ms-2" style={{fontSize: '0.65rem'}}>NEW</Badge>}
                                    </Col>
                                </Row>
                                <Row>
                                    <Col>
                                        {p.projectNumber && <Badge bg="info" className="text-dark" style={{fontSize: "0.65rem"}} title="Project Number">{p.projectNumber}</Badge>}
                                    </Col>
                                    <Col xs="auto">
                                        <Badge bg={p.percent === "100%" ? "success" : p.percent === "0%" ? "danger" : "warning"} pill className="me-2" style={{minWidth: "45px"}} title="Percent Complete">
                                            {p.percent || "0%"}
                                        </Badge>
                                        <Button variant="light" size="sm" className="text-primary p-1 lh-1" onClick={() => handleJump(p.rowIndex)}>
                                            <FontAwesomeIcon icon={faLocationArrow} />
                                        </Button>
                                    </Col>
                                </Row>                            
                            <div className="mt-2 small text-muted">
                                {/* Task Summary */}
                                {/* CLICKABLE TASKS BUTTON */}
                                <div 
                                    className="d-flex justify-content-between mb-1 text-dark p-1 rounded border border-light" 
                                    style={{cursor: "pointer", backgroundColor: "#f8f9fa", transition: "all 0.2s"}}
                                    onClick={() => setSelectedProject(p)}
                                    onMouseOver={(e) => {
                                        e.currentTarget.style.background = "#e2e6ea";
                                        e.currentTarget.style.borderColor = "#dee2e6";
                                    }}
                                    onMouseOut={(e) => {
                                        e.currentTarget.style.background = "#f8f9fa";
                                        e.currentTarget.style.borderColor = "#f8f9fa";
                                    }}
                                    title="Click to Manage Tasks"
                                >
                                    <span className="fw-bold" style={{fontSize: "0.9em"}}>
                                        <FontAwesomeIcon icon={faListCheck} className="me-2 text-primary" />  
                                        Tasks: {p.completedTasks}/{p.totalTasks}
                                    </span>
                                    <span className="text-primary"><FontAwesomeIcon icon={faChevronRight} style={{fontSize: "0.8em"}} /></span>
                                </div>

                                <div className="d-flex justify-content-between mb-1 px-1">
                                    <span><FontAwesomeIcon icon={faUser} className="me-2 text-secondary" style={{width: "14px", textAlign: "center"}} /> {p.lead || "Unassigned"}</span>
                                </div>
                                <div className="d-flex justify-content-between border-top pt-1 mt-1 px-1">
                                    <span className="d-flex align-items-center">
                                        <FontAwesomeIcon icon={faCalendarDays} className="me-2 text-secondary" style={{width: "14px", textAlign: "center"}} />
                                        {p.start === "TBD" || p.start === "" ? "TBD" : p.start}
                                    </span>
                                    {p.start !== "TBD" && p.start !== "" && (
                                        <>
                                            <span className="mx-1 text-muted"><FontAwesomeIcon icon={faArrowRight} style={{fontSize: "0.7rem"}} /></span>
                                            <span>{p.end}</span>
                                        </>
                                    )}
                                </div>
                            </div>
                        </Card.Body>
                    </Card>
                    );
                })}
            </div>
        </div>
    );
};

export default ProjectList;
