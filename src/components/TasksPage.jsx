/* global React, ReactBootstrap, Excel */
import React, { useState, useEffect, useRef } from 'react';
import { GanttLogic } from '../utils/ganttLogic';
import { ChangelogLogic } from '../utils/changelogLogic';
import { IdentityLogic } from '../utils/identityLogic';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { 
    faLocationArrow, 
    faPencil, 
    faTrash, 
    faUser, 
    faCalendarDays, 
    faArrowRight, 
    faSortAmountDown, 
    faSortAmountUp, 
    faSliders, 
    faSearch, 
    faFolder, 
    faClock, 
    faCheckCircle, 
    faPlay, 
    faClipboardList, 
    faSyncAlt,
    faExclamationTriangle
} from '@fortawesome/free-solid-svg-icons';
import { Button, Card, Badge, Spinner, Modal, ButtonGroup, Form, Row, Col, Alert, Collapse } from 'react-bootstrap';

const TasksPage = ({ refreshTrigger }) => {
    const [tasks, setTasks] = useState([]);
    const [projects, setProjects] = useState([]); // Unique list of projects for filtering
    const [teamMembers, setTeamMembers] = useState([]);
    const [isLoading, setIsLoading] = useState(true);

    // MODAL STATE
    const [showDelete, setShowDelete] = useState(false);
    const [showForm, setShowForm] = useState(false);
    const [activeTask, setActiveTask] = useState(null);
    const [formData, setFormData] = useState({ name: "", lead: "", start: "", end: "", percent: "0%" });
    const [status, setStatus] = useState("");

    // FILTER & SORT STATE
    const [filters, setFilters] = useState({ 
        location: "", 
        projectId: "", 
        lead: "", 
        status: "", 
        search: "" 
    });
    const [sortConfig, setSortConfig] = useState({ key: "end", direction: "asc" });
    const [showControls, setShowControls] = useState(false);

    const currentUser = IdentityLogic.getIdentity();

    // Height calculations
    const listContainerRef = useRef(null);
    const [listHeight, setListHeight] = useState("500px");

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
    }, []);

    // --- FETCH ALL PROJECTS & TASKS ---
    const fetchAllTasks = async () => {
        setIsLoading(true);
        try {
            await Excel.run(async (context) => {
                const sheets = ["Houston", "Dallas"];
                const allTasks = [];
                const allProjects = [];

                // 1. Fetch Team Members
                const teamSheet = context.workbook.worksheets.getItemOrNullObject("Team");
                teamSheet.load("isNullObject");
                await context.sync();

                const teamMap = {};
                const members = [];
                if (!teamSheet.isNullObject) {
                    const teamRange = teamSheet.getUsedRange();
                    teamRange.load("text");
                    await context.sync();
                    const teamRows = teamRange.text;
                    for (let i = 1; i < teamRows.length; i++) {
                        const firstName = teamRows[i]?.[0]?.trim() || "";
                        const lastName = teamRows[i]?.[1]?.trim() || "";
                        if (firstName) {
                            const full = `${firstName} ${lastName}`.trim();
                            teamMap[firstName.toLowerCase()] = full;
                            members.push({ first: firstName, full: full });
                        }
                    }
                }
                setTeamMembers(members);

                // 2. Fetch Tasks from Gantt sheets
                for (const sheetName of sheets) {
                    const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                    const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                    footerRange.load(["rowIndex", "isNullObject"]);
                    sheet.load("isNullObject");
                    await context.sync();

                    if (sheet.isNullObject || footerRange.isNullObject) continue;

                    const dataStartIndex = 7;
                    const rowCount = footerRange.rowIndex - dataStartIndex;
                    if (rowCount <= 0) continue;

                    // Find Project Number Column
                    const headerRow = sheet.getRange("7:7");
                    const projNumHeader = headerRow.find("Project Number", { completeMatch: true, matchCase: false });
                    projNumHeader.load(["columnIndex", "isNullObject"]);
                    await context.sync();

                    let projNumIdx = 3;
                    if (!projNumHeader.isNullObject) {
                        projNumIdx = projNumHeader.columnIndex;
                    }

                    const colCountToFetch = Math.max(8, projNumIdx + 1);
                    const range = sheet.getRangeByIndexes(dataStartIndex, 0, rowCount, colCountToFetch);
                    range.load("text");
                    await context.sync();

                    const rawRows = range.text;
                    let currentProject = null;

                    rawRows.forEach((row, index) => {
                        if (!row[1] || row[1] === "") return;
                        const idStr = row[0].toString();
                        const idNum = parseFloat(idStr);
                        const currentRowIndex = dataStartIndex + index;

                        if (!isNaN(idNum) && Number.isInteger(idNum)) {
                            // It is a PROJECT
                            const rawLead = row[2]?.trim() || "";
                            const fullLeadName = teamMap[rawLead.toLowerCase()] || rawLead;
                            currentProject = {
                                id: idStr,
                                name: row[1],
                                projectNumber: row[projNumIdx],
                                lead: fullLeadName,
                                location: sheetName,
                                rowIndex: currentRowIndex
                            };
                            allProjects.push(currentProject);
                        } else if (!isNaN(idNum) && !Number.isInteger(idNum)) {
                            // It is a TASK
                            const cleanName = row[1].toString().replace(/^[↑\s]+/, '');
                            const dotCount = (idStr.match(/\./g) || []).length;
                            const depth = Math.max(0, dotCount - 1);

                            allTasks.push({
                                id: idStr,
                                name: cleanName,
                                lead: row[2], // raw Lead (typically first name)
                                start: row[4],
                                end: row[5],
                                percent: row[7],
                                depth: depth,
                                rowIndex: currentRowIndex,
                                location: sheetName,
                                project: currentProject ? { ...currentProject } : { id: Math.floor(idNum).toString(), name: "Unknown Project", location: sheetName }
                            });
                        }
                    });
                }

                setTasks(allTasks);
                setProjects(allProjects);
            });
        } catch (error) {
            console.error("Fetch all tasks error:", error);
        } finally {
            setIsLoading(false);
        }
    };

    useEffect(() => {
        fetchAllTasks();
    }, [refreshTrigger]);

    // --- LOCATE TASK IN EXCEL ---
    const handleJump = async (task) => {
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(task.location);
                sheet.activate();
                const range = sheet.getRangeByIndexes(task.rowIndex, 0, 1, 1).getEntireRow();
                range.select();
                await context.sync();
            });
        } catch (error) {
            console.error("Excel Jump Error:", error);
        }
    };

    // --- EDIT TASK MODAL ---
    const openEditModal = (task) => {
        setActiveTask(task);
        const formatDateInput = (dateStr) => {
            if (!dateStr || dateStr === "TBD") return "";
            const d = new Date(dateStr);
            return !isNaN(d) ? d.toISOString().split('T')[0] : "";
        };
        setFormData({
            name: task.name,
            lead: task.lead,
            start: formatDateInput(task.start),
            end: formatDateInput(task.end),
            percent: task.percent
        });
        setStatus("");
        setShowForm(true);
    };

    const handleSave = async () => {
        if (!formData.name) {
            setStatus("Task Name is required.");
            return;
        }

        setStatus("");
        const loaderMessage = `Updating Task Row ${activeTask.rowIndex + 1} in ${activeTask.location}...`;
        if (window.GlobalLoader) window.GlobalLoader.show(loaderMessage);

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(activeTask.location);
                const rowIndex = activeTask.rowIndex;

                sheet.getCell(rowIndex, 1).values = [[formData.name]];
                if (formData.lead) sheet.getCell(rowIndex, 2).values = [[formData.lead]];
                if (formData.start) sheet.getCell(rowIndex, 4).values = [[formData.start]];
                
                if (formData.start && formData.end) {
                    const s = new Date(formData.start);
                    const e = new Date(formData.end);
                    const diff = Math.ceil(Math.abs(e - s) / (1000 * 60 * 60 * 24)) + 1;
                    sheet.getCell(rowIndex, 6).values = [[diff]];
                }
                sheet.getCell(rowIndex, 7).values = [[formData.percent]];

                await ChangelogLogic.logChange(context, `Edited Task: "${formData.name}" (ID: ${activeTask.id}) in Project ${activeTask.project.id} via Tasks Page`);

                // Update averages
                await GanttLogic.updateProjectAverages(context, activeTask.location);
                await context.sync();
            });

            setShowForm(false);
            if (window.GlobalToast) window.GlobalToast.success("Task updated successfully!");
            await fetchAllTasks();
        } catch (err) {
            console.error(err);
            setStatus("Error: " + err.message);
        } finally {
            if (window.GlobalLoader) window.GlobalLoader.hide();
        }
    };

    // --- DELETE TASK ---
    const handleDelete = async () => {
        if (!activeTask) return;

        if (window.GlobalLoader) {
            window.GlobalLoader.show(`Deleting Task Row ${activeTask.rowIndex + 1} from ${activeTask.location}...`);
        }

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(activeTask.location);
                await ChangelogLogic.logChange(context, `Deleted Task: "${activeTask.name}" from Project ${activeTask.project.id} via Tasks Page`);
                const range = sheet.getRangeByIndexes(activeTask.rowIndex, 0, 1, 1).getEntireRow();
                range.delete(Excel.DeleteShiftDirection.up);

                // Update averages
                await GanttLogic.updateProjectAverages(context, activeTask.location);
                await context.sync();
            });

            setShowDelete(false);
            setActiveTask(null);
            if (window.GlobalToast) window.GlobalToast.error("Task deleted successfully!");
            await fetchAllTasks();
        } catch (error) {
            console.error(error);
        } finally {
            if (window.GlobalLoader) window.GlobalLoader.hide();
        }
    };

    // --- PROCESS FILTERS & SORTS ---
    const getProcessedTasks = () => {
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        let filtered = tasks.filter(t => {
            // 1. Search Query
            const query = filters.search.toLowerCase().trim();
            const matchesSearch = !query || 
                t.name.toLowerCase().includes(query) || 
                t.id.includes(query) || 
                t.project.name.toLowerCase().includes(query) ||
                (t.project.projectNumber && t.project.projectNumber.toLowerCase().includes(query));

            // 2. Location
            const matchesLocation = !filters.location || t.location === filters.location;

            // 3. Project
            const matchesProject = !filters.projectId || t.project.id === filters.projectId;

            // 4. Lead/Assignee
            const matchesLead = !filters.lead || t.lead?.toLowerCase() === filters.lead?.toLowerCase();

            // 5. Status
            let matchesStatus = true;
            if (filters.status) {
                const isCompleted = t.percent === "100%";
                const isNotStarted = t.percent === "0%" || !t.percent;
                const isOverdue = !isCompleted && t.end !== "TBD" && t.end !== "" && new Date(t.end) < today;
                
                if (filters.status === "completed") {
                    matchesStatus = isCompleted;
                } else if (filters.status === "notstarted") {
                    matchesStatus = isNotStarted;
                } else if (filters.status === "active") {
                    matchesStatus = !isCompleted && !isNotStarted && !isOverdue;
                } else if (filters.status === "overdue") {
                    matchesStatus = isOverdue;
                } else if (filters.status === "all") {
                    matchesStatus = true;
                }
            } else {
                // By default (empty string status filter), hide completed tasks
                matchesStatus = t.percent !== "100%";
            }

            return matchesSearch && matchesLocation && matchesProject && matchesLead && matchesStatus;
        });

        // Sorting
        return filtered.sort((a, b) => {
            let valA = a[sortConfig.key];
            let valB = b[sortConfig.key];

            if (sortConfig.key === 'id') {
                const aParts = valA.split('.').map(Number);
                const bParts = valB.split('.').map(Number);
                for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
                    const numA = aParts[i] || 0;
                    const numB = bParts[i] || 0;
                    if (numA !== numB) return sortConfig.direction === 'asc' ? numA - numB : numB - numA;
                }
                return 0;
            } else if (sortConfig.key === 'percent') {
                valA = parseInt(valA?.replace('%', '') || '0');
                valB = parseInt(valB?.replace('%', '') || '0');
            } else if (sortConfig.key === 'start' || sortConfig.key === 'end') {
                const parseDate = (dStr) => {
                    if (!dStr || dStr === "TBD") return null;
                    const d = new Date(dStr);
                    return isNaN(d) ? null : d.getTime();
                };
                const timeA = parseDate(valA);
                const timeB = parseDate(valB);

                if (timeA === null && timeB === null) return 0;
                if (timeA === null) return 1; // Keep TBDs at the bottom
                if (timeB === null) return -1;

                return sortConfig.direction === 'asc' ? timeA - timeB : timeB - timeA;
            } else if (sortConfig.key === 'project.name') {
                valA = a.project.name.toLowerCase();
                valB = b.project.name.toLowerCase();
            } else {
                valA = (valA || "").toString().toLowerCase();
                valB = (valB || "").toString().toLowerCase();
            }

            if (valA < valB) return sortConfig.direction === 'asc' ? -1 : 1;
            if (valA > valB) return sortConfig.direction === 'asc' ? 1 : -1;
            return 0;
        });
    };

    const processedTasks = getProcessedTasks();

    // Unique project lists for dropdown (filtered by selected location)
    const filteredProjects = projects.filter(p => !filters.location || p.location === filters.location);

    return (
        <div className="mt-2">
            {/* TITLE HEADER */}
            <div className="d-flex justify-content-between align-items-center mb-3 pb-2 border-bottom">
                <h5 className="m-0 fw-bold text-primary">
                    <FontAwesomeIcon icon={faClipboardList} className="me-2" />
                    All Project Tasks
                </h5>
                <div className="d-flex align-items-center">
                    {currentUser && (
                        <Button 
                            variant="link" 
                            size="sm" 
                            className={`text-decoration-none p-0 me-3 ${filters.lead?.toLowerCase() === currentUser.toLowerCase() ? "text-primary fw-bold" : "text-muted"}`} 
                            onClick={() => {
                                const isCurrentlyMine = filters.lead?.toLowerCase() === currentUser.toLowerCase();
                                setFilters({
                                    ...filters,
                                    lead: isCurrentlyMine ? "" : currentUser
                                });
                            }} 
                            title="Toggle My Tasks"
                        >
                            <FontAwesomeIcon icon={faUser} className="me-1" />
                            My Tasks
                        </Button>
                    )}
                    <Button variant="link" size="sm" className="text-decoration-none p-0 me-3 text-muted" onClick={() => setShowControls(!showControls)} title="Filters & Sort">
                        <FontAwesomeIcon icon={faSliders} className={showControls ? "text-primary" : ""} />
                    </Button>
                    <Button variant="link" size="sm" className="text-decoration-none p-0" onClick={fetchAllTasks}>
                        {isLoading ? <Spinner animation="border" size="sm" /> : <><FontAwesomeIcon icon={faSyncAlt} className="me-1" /> Refresh</>}
                    </Button>
                </div>
            </div>

            {/* SEARCH INPUT */}
            <Form.Group className="mb-2 position-relative">
                <Form.Control 
                    type="text" 
                    placeholder="Search by task name, project name, ID..." 
                    value={filters.search}
                    onChange={(e) => setFilters({ ...filters, search: e.target.value })}
                    size="sm"
                    className="pe-4"
                />
                <FontAwesomeIcon 
                    icon={faSearch} 
                    className="position-absolute text-muted" 
                    style={{ right: '12px', top: '10px', fontSize: '0.8rem' }}
                />
            </Form.Group>

            {/* CONTROLS PANEL */}
            <Collapse in={showControls}>
                <div>
                    <div className="bg-light p-2 rounded mb-3 border shadow-sm small">
                        <Row className="g-2 mb-2">
                            <Col xs={6}>
                                <Form.Label className="small fw-bold text-muted mb-1 text-uppercase" style={{ fontSize: '0.65rem' }}>Location</Form.Label>
                                <Form.Select 
                                    size="sm" 
                                    value={filters.location} 
                                    onChange={(e) => setFilters({ ...filters, location: e.target.value, projectId: "" })}
                                >
                                    <option value="">All Locations</option>
                                    <option value="Houston">Houston</option>
                                    <option value="Dallas">Dallas</option>
                                </Form.Select>
                            </Col>
                            <Col xs={6}>
                                <Form.Label className="small fw-bold text-muted mb-1 text-uppercase" style={{ fontSize: '0.65rem' }}>Project</Form.Label>
                                <Form.Select 
                                    size="sm" 
                                    value={filters.projectId} 
                                    onChange={(e) => setFilters({ ...filters, projectId: e.target.value })}
                                >
                                    <option value="">All Projects</option>
                                    {filteredProjects.map((p, idx) => (
                                        <option key={idx} value={p.id}>
                                            #{p.id} {p.name} ({p.location})
                                        </option>
                                    ))}
                                </Form.Select>
                            </Col>
                        </Row>
                        <Row className="g-2">
                            <Col xs={4}>
                                <Form.Label className="small fw-bold text-muted mb-1 text-uppercase" style={{ fontSize: '0.65rem' }}>Assignee</Form.Label>
                                <Form.Select 
                                    size="sm" 
                                    value={filters.lead} 
                                    onChange={(e) => setFilters({ ...filters, lead: e.target.value })}
                                >
                                    <option value="">All Leads</option>
                                    {teamMembers.map((m, idx) => <option key={idx} value={m.first}>{m.full}</option>)}
                                </Form.Select>
                            </Col>
                            <Col xs={4}>
                                <Form.Label className="small fw-bold text-muted mb-1 text-uppercase" style={{ fontSize: '0.65rem' }}>Status</Form.Label>
                                <Form.Select 
                                    size="sm" 
                                    value={filters.status} 
                                    onChange={(e) => setFilters({ ...filters, status: e.target.value })}
                                >
                                    <option value="">Hide Completed (Default)</option>
                                    <option value="all">Show All Statuses</option>
                                    <option value="overdue">⚠️ Overdue</option>
                                    <option value="active">▶️ Active</option>
                                    <option value="notstarted">⏹️ Not Started</option>
                                    <option value="completed">✅ Completed</option>
                                </Form.Select>
                            </Col>
                            <Col xs={4}>
                                <Form.Label className="small fw-bold text-muted mb-1 text-uppercase" style={{ fontSize: '0.65rem' }}>Sort By</Form.Label>
                                <div className="d-flex">
                                    <Form.Select 
                                        size="sm" 
                                        className="me-1"
                                        value={sortConfig.key} 
                                        onChange={(e) => setSortConfig({ ...sortConfig, key: e.target.value })}
                                    >
                                        <option value="id">Task #</option>
                                        <option value="name">Task Name</option>
                                        <option value="project.name">Project Name</option>
                                        <option value="start">Start Date</option>
                                        <option value="end">End Date</option>
                                        <option value="percent">% Complete</option>
                                    </Form.Select>
                                    <Button 
                                        variant="outline-secondary" 
                                        size="sm" 
                                        onClick={() => setSortConfig({ ...sortConfig, direction: sortConfig.direction === 'asc' ? 'desc' : 'asc' })}
                                    >
                                        <FontAwesomeIcon icon={sortConfig.direction === 'asc' ? faSortAmountUp : faSortAmountDown} />
                                    </Button>
                                </div>
                            </Col>
                        </Row>
                    </div>
                </div>
            </Collapse>

            {/* LIST OF TASKS */}
            {isLoading ? (
                <div className="text-center py-5"><Spinner animation="border" size="sm" variant="primary" /></div>
            ) : processedTasks.length === 0 ? (
                <div className="text-center text-muted small mt-5 py-4">
                    <FontAwesomeIcon icon={faClipboardList} size="2x" className="mb-2 text-secondary opacity-50" /><br/>
                    No tasks found matching these criteria.<br/>
                    <Button variant="link" size="sm" className="p-0 mt-2 text-decoration-none fw-bold" onClick={() => setFilters({ location: "", projectId: "", lead: "", status: "", search: "" })}>
                        Clear all filters
                    </Button>
                </div>
            ) : (
                <div 
                    ref={listContainerRef} 
                    style={{ maxHeight: listHeight, overflowY: "auto", paddingRight: "4px" }}
                >
                    {processedTasks.map((t, idx) => {
                        const isCompleted = t.percent === "100%";
                        const isNotStarted = t.percent === "0%" || !t.percent;
                        const today = new Date();
                        today.setHours(0, 0, 0, 0);
                        const isOverdue = !isCompleted && t.end !== "TBD" && t.end !== "" && new Date(t.end) < today;

                        let statusIcon = faPlay;
                        let statusColor = "warning";
                        if (isCompleted) {
                            statusIcon = faCheckCircle;
                            statusColor = "success";
                        } else if (isNotStarted) {
                            statusIcon = faClock;
                            statusColor = "secondary";
                        } else if (isOverdue) {
                            statusIcon = faExclamationTriangle;
                            statusColor = "danger";
                        }

                        return (
                            <Card 
                                key={idx} 
                                className={`mb-2 shadow-sm border-0 ${isOverdue ? 'border-start border-danger border-3' : ''}`}
                                style={{ marginLeft: `${t.depth * 8}px`, borderLeft: t.depth > 0 ? "2px solid #e9ecef" : "none" }}
                            >
                                <Card.Body className="p-2">
                                    {/* Project info header */}
                                    <div className="d-flex justify-content-between align-items-center mb-1 text-muted border-bottom pb-1" style={{ fontSize: '0.7rem' }}>
                                        <span>
                                            <FontAwesomeIcon icon={faFolder} className="me-1 opacity-75 text-primary" />
                                            <strong>{t.location.toUpperCase()}</strong> | Project #{t.project.id} {t.project.name}
                                        </span>
                                        <Badge bg="light" text="dark" className="border">{t.id}</Badge>
                                    </div>

                                    {/* Task Name and Buttons */}
                                    <div className="d-flex justify-content-between align-items-start">
                                        <div className="fw-bold text-dark text-truncate" title={t.name} style={{ fontSize: '0.85rem' }}>
                                            {t.name}
                                        </div>
                                        <ButtonGroup size="sm" className="ms-2 flex-shrink-0">
                                            <Button variant="light" className="px-2 text-primary" onClick={() => handleJump(t)} title="Locate in Excel">
                                                <FontAwesomeIcon icon={faLocationArrow} style={{ fontSize: '0.7rem' }} />
                                            </Button>
                                            <Button variant="light" className="px-2 text-secondary" onClick={() => openEditModal(t)} title="Edit Task">
                                                <FontAwesomeIcon icon={faPencil} style={{ fontSize: '0.7rem' }} />
                                            </Button>
                                            <Button variant="light" className="px-2 text-danger" onClick={() => { setActiveTask(t); setShowDelete(true); }} title="Delete Task">
                                                <FontAwesomeIcon icon={faTrash} style={{ fontSize: '0.7rem' }} />
                                            </Button>
                                        </ButtonGroup>
                                    </div>

                                    {/* Task details (assignee, dates, complete) */}
                                    <div className="mt-2 small text-muted" style={{ fontSize: '0.75rem' }}>
                                        <div className="d-flex justify-content-between align-items-center">
                                            <span>
                                                <FontAwesomeIcon icon={faUser} className="me-1 opacity-50" />
                                                {(() => {
                                                    if (!t.lead) return "Unassigned";
                                                    const member = teamMembers.find(m => m.first === t.lead);
                                                    return member ? member.full : t.lead;
                                                })()}
                                            </span>
                                            <Badge bg={statusColor} text={statusColor === "light" ? "dark" : "white"} className="fw-normal">
                                                <FontAwesomeIcon icon={statusIcon} className="me-1" />
                                                {t.percent || "0%"}
                                            </Badge>
                                        </div>
                                        <div className="d-flex justify-content-between border-top pt-1 mt-1">
                                            <span className="d-flex align-items-center">
                                                <FontAwesomeIcon icon={faCalendarDays} className="me-1 opacity-75" />
                                                {t.start === "TBD" || t.start === "" ? "TBD" : t.start}
                                            </span>
                                            {t.start !== "TBD" && t.start !== "" && (
                                                <>
                                                    <span className="mx-1 text-muted"><FontAwesomeIcon icon={faArrowRight} style={{ fontSize: '0.65rem' }} /></span>
                                                    <span>{t.end}</span>
                                                </>
                                            )}
                                        </div>
                                    </div>
                                </Card.Body>
                            </Card>
                        );
                    })}
                </div>
            )}

            {/* EDIT FORM MODAL */}
            <Modal show={showForm} onHide={() => setShowForm(false)} centered>
                <Modal.Header closeButton className="py-2 bg-light">
                    <Modal.Title style={{ fontSize: '1rem' }} className="fw-bold">Edit Task</Modal.Title>
                </Modal.Header>
                <Modal.Body className="p-3">
                    <Form.Group className="mb-2">
                        <Form.Label className="small fw-bold text-muted">TASK NAME</Form.Label>
                        <Form.Control size="sm" type="text" value={formData.name} onChange={(e) => setFormData({ ...formData, name: e.target.value })} />
                    </Form.Group>
                    
                    <Form.Group className="mb-2">
                        <Form.Label className="small fw-bold text-muted">ASSIGN LEAD</Form.Label>
                        <Form.Select size="sm" value={formData.lead} onChange={(e) => setFormData({ ...formData, lead: e.target.value })}>
                            <option value="">Unassigned</option>
                            {teamMembers.map((m, idx) => <option key={idx} value={m.first}>{m.full}</option>)}
                        </Form.Select>
                    </Form.Group>

                    <Row className="g-2 mb-2">
                        <Col xs={6}>
                            <Form.Group>
                                <Form.Label className="small fw-bold text-muted">START</Form.Label>
                                <Form.Control 
                                    size="sm" 
                                    type="date" 
                                    value={formData.start} 
                                    onChange={(e) => {
                                        const newStart = e.target.value;
                                        setFormData(prev => {
                                            const nextData = { ...prev, start: newStart };
                                            if (newStart && (!prev.end || prev.end < newStart)) {
                                                nextData.end = newStart;
                                            }
                                            return nextData;
                                        });
                                    }} 
                                />
                            </Form.Group>
                        </Col>
                        <Col xs={6}>
                            <Form.Group>
                                <Form.Label className="small fw-bold text-muted">END (EST)</Form.Label>
                                <Form.Control size="sm" type="date" value={formData.end} onChange={(e) => setFormData({ ...formData, end: e.target.value })} />
                            </Form.Group>
                        </Col>
                    </Row>

                    <Form.Group className="mb-2">
                        <div className="d-flex justify-content-between align-items-center mb-1">
                            <Form.Label className="small fw-bold text-muted m-0">% COMPLETE</Form.Label>
                            <Badge bg="primary" style={{ fontSize: '0.7rem' }}>{formData.percent || "0%"}</Badge>
                        </div>
                        <input 
                            type="range"
                            min={0}
                            max={100}
                            step={1}
                            value={parseInt(formData.percent) || 0}
                            onChange={(e) => setFormData({ ...formData, percent: e.target.value + "%" })}
                            className="w-100 mt-2"
                            style={{ accentColor: '#0d6efd', cursor: 'pointer' }}
                        />
                    </Form.Group>

                    {status && <div className="text-danger small mt-2">{status}</div>}
                </Modal.Body>
                <Modal.Footer className="py-1 bg-light">
                    <Button variant="secondary" size="sm" onClick={() => setShowForm(false)}>Cancel</Button>
                    <Button variant="primary" size="sm" onClick={handleSave}>Save Task</Button>
                </Modal.Footer>
            </Modal>

            {/* DELETE MODAL */}
            <Modal show={showDelete} onHide={() => setShowDelete(false)} centered size="sm">
                <Modal.Header closeButton className="py-2 bg-light">
                    <Modal.Title style={{ fontSize: '1rem' }} className="fw-bold text-danger">Delete Task?</Modal.Title>
                </Modal.Header>
                <Modal.Body className="small text-center py-4">
                    Are you sure you want to delete <br/>
                    <strong className="text-dark">{activeTask?.name}</strong>?
                </Modal.Body>
                <Modal.Footer className="py-1 bg-light">
                    <Button variant="secondary" size="sm" onClick={() => setShowDelete(false)}>Cancel</Button>
                    <Button variant="danger" size="sm" onClick={handleDelete}>Yes, Delete</Button>
                </Modal.Footer>
            </Modal>
        </div>
    );
};

export default TasksPage;
