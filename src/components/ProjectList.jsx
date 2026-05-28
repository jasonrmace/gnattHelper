/* global React, ReactBootstrap, Excel */
import ProjectTasks from './ProjectTasks';

import React, { useState, useEffect, useRef } from 'react';
import { Button, Card, Badge, Spinner, Row, Col } from 'react-bootstrap';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faLocationArrow, faListCheck, faChevronRight, faUser, faCalendarDays, faArrowRight, faSyncAlt } from '@fortawesome/free-solid-svg-icons';

const ProjectList = ({ refreshTrigger, highlightId }) => {
    // --- STATE ---
    const [projects, setProjects] = useState([]);
    const [isFetching, setIsFetching] = useState(false);
    
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
                const sheet = context.workbook.worksheets.getItem("GanttChart");
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
                const teamSheet = context.workbook.worksheets.getItem("Team");
                const teamRange = teamSheet.getUsedRange();
                teamRange.load("text");

                const sheet = context.workbook.worksheets.getItem("GanttChart");
                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                footerRange.load("rowIndex");
                await context.sync();

                const teamMap = {};
                const teamRows = teamRange.text;
                for (let i = 1; i < teamRows.length; i++) {
                    const firstName = teamRows[i]?.[0]?.trim() || "";
                    const lastName = teamRows[i]?.[1]?.trim() || "";
                    if (firstName) teamMap[firstName.toLowerCase()] = `${firstName} ${lastName}`.trim();
                }

                const dataStartIndex = 7; 
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
                <h6 className="m-0 fw-bold text-primary">Active Projects ({projects.length})</h6>
                <Button variant="link" size="sm" className="text-decoration-none p-0" onClick={fetchProjects}>
                    {isFetching ? <Spinner animation="border" size="sm" /> : <><FontAwesomeIcon icon={faSyncAlt} className="me-1" /> Refresh</>}
                </Button>
            </div>

            {projects.length === 0 && !isFetching && (
                <div className="text-center text-muted small mt-2">No projects found.</div>
            )}

            <div 
                ref={listContainerRef}
                style={{ maxHeight: listHeight, overflowY: "auto", transition: "max-height 0.1s ease-out" }}
            >
                {projects.map((p, index) => {
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
