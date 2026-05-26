/* global React, ReactBootstrap, Excel */

const { useState, useEffect } = React;
const { Button, Card, Badge, Spinner, Modal, ButtonGroup, Form, Row, Col, Alert } = window.ReactBootstrap || {};

const ProjectTasks = ({ project, onBack }) => {
    const [tasks, setTasks] = useState([]);
    const [teamMembers, setTeamMembers] = useState([]); // Stores { first, full }
    const [isLoading, setIsLoading] = useState(false);
    const [showDelete, setShowDelete] = useState(false);
    const [showForm, setShowForm] = useState(false);
    const [activeTask, setActiveTask] = useState(null);
    const [parentTask, setParentTask] = useState(null);
    const [formMode, setFormMode] = useState("add");
    const [formData, setFormData] = useState({ name: "", lead: "", start: "", end: "", percent: "0%" });
    const [status, setStatus] = useState("");

    // --- 1. INITIAL LOADING ---
    useEffect(() => {
        const fetchTeam = async () => {
            try {
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getItem("Team");
                    const range = sheet.getUsedRange(); 
                    // Assuming Team Table: Col A = First Name, Col B = Last Name
                    range.load("text");
                    await context.sync();
                    
                    const rows = range.text;
                    const members = [];
                    // Start at i=1 to skip header
                    for (let i = 1; i < rows.length; i++) {
                        const first = rows[i][0]?.trim();
                        const last = rows[i][1]?.trim();
                        if (first) {
                            members.push({ 
                                first: first, 
                                full: last ? `${first} ${last}` : first 
                            });
                        }
                    }
                    setTeamMembers(members);
                });
            } catch (e) {
                console.error("Team load error", e);
            }
        };
        fetchTeam();
    }, []);

    const fetchTasks = async () => {
        setIsLoading(true);
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");
                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                footerRange.load("rowIndex");
                await context.sync();

                const dataStartIndex = 7; 
                const rowCount = footerRange.rowIndex - dataStartIndex;

                if (rowCount <= 0) {
                    setTasks([]);
                    return;
                }

                const range = sheet.getRangeByIndexes(dataStartIndex, 0, rowCount, 8);
                range.load("text");
                await context.sync();

                const rawRows = range.text;
                const projectTasks = [];
                const parentIdPrefix = `${project.id}.`;

                rawRows.forEach((row, index) => {
                    const idStr = row[0].toString();
                    
                    // Check if it belongs to this project (e.g. "4.1", "4.2") but IS NOT the header "4"
                    if (idStr.startsWith(parentIdPrefix) && idStr !== project.id.toString()) {
                        const dotCount = (idStr.match(/\./g) || []).length;
                        const depth = Math.max(0, dotCount - 1); // Indentation level
                        const cleanName = row[1].toString().replace(/^[↑\s]+/, '');

                        projectTasks.push({
                            id: idStr,
                            rowIndex: dataStartIndex + index,
                            name: cleanName,
                            lead: row[2], // This is usually just First Name
                            start: row[4],
                            end: row[5],
                            percent: row[7],
                            depth: depth
                        });
                    }
                });
                setTasks(projectTasks);
            });
        } catch (error) {
            console.error("Fetch tasks error:", error);
        } finally {
            setIsLoading(false);
        }
    };

    useEffect(() => {
        fetchTasks();
    }, [project]);

    // --- 2. ACTIONS ---
    const openAddModal = () => {
        setFormMode("add");
        setParentTask(null);
        setFormData({ name: "", lead: "", start: "", end: "", percent: "0%" });
        setStatus("");
        setShowForm(true);
    };

    const openSubTaskModal = (parent) => {
        setFormMode("sub");
        setParentTask(parent);
        setFormData({ name: "", lead: "", start: "", end: "", percent: "0%" });
        setStatus("");
        setShowForm(true);
    };

    const openEditModal = (task) => {
        setFormMode("edit");
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

    // --- 3. SAVE LOGIC (WITH GANTT LOGIC TRIGGER) ---
    const handleSave = async () => {
        if (!formData.name) {
            setStatus("Task Name is required.");
            return;
        }
        setStatus("Processing...");

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");

                if (formMode === "edit") {
                    const rowIndex = activeTask.rowIndex;
                    sheet.getCell(rowIndex, 1).values = [[formData.name]];
                    if (formData.lead) sheet.getCell(rowIndex, 2).values = [[formData.lead]]; // Saves First Name
                    if (formData.start) sheet.getCell(rowIndex, 4).values = [[formData.start]];
                    
                    if (formData.start && formData.end) {
                        const s = new Date(formData.start);
                        const e = new Date(formData.end);
                        const diff = Math.ceil(Math.abs(e - s) / (1000 * 60 * 60 * 24)) + 1;
                        sheet.getCell(rowIndex, 6).values = [[diff]];
                    }
                    sheet.getCell(rowIndex, 7).values = [[formData.percent]];
                } else {
                    // ADD NEW TASK LOGIC
                    let templateName = "Level2Task";
                    let insertRowIndex = -1;

                    if (formMode === "sub" && parentTask) {
                        if (parentTask.depth === 0) templateName = "Level3Task";
                        else if (parentTask.depth >= 1) templateName = "Level4Task";
                        insertRowIndex = parentTask.rowIndex + 1;
                    } else {
                        if (tasks.length > 0) {
                            const lastTask = tasks[tasks.length - 1];
                            insertRowIndex = lastTask.rowIndex + 1;
                        } else {
                            insertRowIndex = project.rowIndex + 1;
                        }
                    }

                    let namedItem = sheet.names.getItemOrNullObject(templateName);
                    namedItem.load("isNullObject");
                    await context.sync();

                    if (namedItem.isNullObject) {
                         namedItem = context.workbook.names.getItemOrNullObject(templateName);
                         namedItem.load("isNullObject");
                         await context.sync();
                    }

                    if (namedItem.isNullObject) throw new Error(`Template '${templateName}' not found.`);

                    const sourceRow = namedItem.getRange().getEntireRow();
                    const insertRange = sheet.getRange(`${insertRowIndex + 1}:${insertRowIndex + 1}`);
                    insertRange.insert(Excel.InsertShiftDirection.down);

                    const newRow = sheet.getRange(`${insertRowIndex + 1}:${insertRowIndex + 1}`);
                    newRow.copyFrom(sourceRow, Excel.RangeCopyType.all);

                    sheet.getCell(insertRowIndex, 1).values = [[formData.name]];
                    if (formData.lead) sheet.getCell(insertRowIndex, 2).values = [[formData.lead]];
                    if (formData.start) sheet.getCell(insertRowIndex, 4).values = [[formData.start]];
                    
                    if (formData.start && formData.end) {
                        const s = new Date(formData.start);
                        const e = new Date(formData.end);
                        const diff = Math.ceil(Math.abs(e - s) / (1000 * 60 * 60 * 24)) + 1;
                        sheet.getCell(insertRowIndex, 6).values = [[diff]];
                    }
                    newRow.select();
                }

                // --- TRIGGER LOGIC ENGINE ---
                if (window.GanttLogic) {
                    await window.GanttLogic.updateProjectAverages(context);
                }
                await context.sync();
                
                setShowForm(false);
                fetchTasks();
            });
        } catch (err) {
            console.error(err);
            setStatus("Error: " + err.message);
        }
    };

    // --- 4. DELETE LOGIC (WITH GANTT LOGIC TRIGGER) ---
    const handleDelete = async () => {
        if (!activeTask) return;
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");
                const range = sheet.getRangeByIndexes(activeTask.rowIndex, 0, 1, 1).getEntireRow();
                range.delete(Excel.DeleteShiftDirection.up);

                // --- TRIGGER LOGIC ENGINE ---
                if (window.GanttLogic) {
                    await window.GanttLogic.updateProjectAverages(context);
                }
                await context.sync();
                
                setShowDelete(false);
                setActiveTask(null);
                fetchTasks();
            });
        } catch (error) {
            console.error(error);
        }
    };

    const handleJump = async (rowIndex) => {
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");
                sheet.activate();
                const range = sheet.getRangeByIndexes(rowIndex, 0, 1, 1).getEntireRow();
                range.select();
                await context.sync();
            });
        } catch (error) {
            console.error(error);
        }
    };

    return (
        <div className="mt-4">
            {/* HEADER */}
            <div className="d-flex align-items-center justify-content-between mb-3 pb-2 border-bottom">
                <div className="d-flex align-items-center">
                    <Button variant="light" size="sm" className="me-2 text-muted" onClick={onBack} title="Back">
                        <i className="fas fa-arrow-left"></i>
                    </Button>
                    <div style={{lineHeight: "1.1"}}>
                        <h6 className="m-0 fw-bold text-primary">{project.name}</h6>
                        <small className="text-muted" style={{fontSize: "0.7rem"}}>Project {project.id} Task Manager</small>
                    </div>
                </div>
                <Button variant="primary" size="sm" onClick={openAddModal}>
                    <i className="fas fa-plus me-1"></i> Add Task
                </Button>
            </div>

            {/* TASK LIST */}
            {isLoading ? (
                <div className="text-center py-5"><Spinner animation="border" size="sm" variant="primary" /></div>
            ) : tasks.length === 0 ? (
                <div className="text-center text-muted small mt-5">
                    <i className="fas fa-clipboard-list fa-2x mb-2 text-secondary opacity-50"></i><br/>
                    No tasks found.
                </div>
            ) : (
                <div style={{ maxHeight: "calc(100vh - 220px)", overflowY: "auto", paddingRight: "4px" }}>
                    {tasks.map((t, idx) => (
                        <Card key={idx} className="mb-2 shadow-sm border-0" style={{ marginLeft: `${t.depth * 24}px`, borderLeft: t.depth > 0 ? "3px solid #e9ecef" : "none" }}>
                            <Card.Body className="p-2">
                                {/* TOP ROW: ID, NAME, BUTTONS */}
                                <div className="d-flex justify-content-between align-items-start">
                                    <div className="d-flex align-items-center" style={{overflow: "hidden"}}>
                                        <Badge bg="secondary" className="me-2" style={{fontSize: "0.65rem", minWidth: "35px"}}>{t.id}</Badge>
                                        <span className="fw-bold text-dark text-truncate" title={t.name} style={{fontSize: "0.85rem"}}>
                                            {t.name}
                                        </span>
                                    </div>
                                    <ButtonGroup size="sm" className="ms-2 flex-shrink-0">
                                        <Button variant="light" className="px-2 text-primary" onClick={() => handleJump(t.rowIndex)} title="Locate">
                                            <i className="fas fa-location-arrow" style={{fontSize: "0.7rem"}}></i>
                                        </Button>
                                        {t.depth < 2 && (
                                            <Button variant="light" className="px-2 text-success" onClick={() => openSubTaskModal(t)} title="Add Sub-Task">
                                                <i className="fas fa-plus" style={{fontSize: "0.7rem"}}></i>
                                            </Button>
                                        )}
                                        <Button variant="light" className="px-2 text-secondary" onClick={() => openEditModal(t)} title="Edit">
                                            <i className="fas fa-pencil" style={{fontSize: "0.7rem"}}></i>
                                        </Button>
                                        <Button variant="light" className="px-2 text-danger" onClick={() => { setActiveTask(t); setShowDelete(true); }} title="Delete">
                                            <i className="fas fa-trash" style={{fontSize: "0.7rem"}}></i>
                                        </Button>
                                    </ButtonGroup>
                                </div>

                                {/* BOTTOM ROW: LEAD (UPDATED), BADGE, DATES */}
                                <div className="mt-2 small text-muted">
                                    <div className="d-flex justify-content-between mb-1 align-items-center">
                                        <span className="d-flex align-items-center">
                                            <i className="fas fa-user me-2 text-secondary opacity-50" style={{width: "14px"}}></i>
                                            
                                            {/* --- UPDATE: FULL NAME LOOKUP --- */}
                                            {(() => {
                                                if (!t.lead) return "-";
                                                // Find match in teamMembers list
                                                const member = teamMembers.find(m => m.first === t.lead);
                                                return member ? member.full : t.lead;
                                            })()}

                                        </span>
                                        <Badge bg={t.percent === "100%" ? "success" : "light"} text={t.percent === "100%" ? "white" : "dark"} className="border fw-normal">
                                            {t.percent || "0%"}
                                        </Badge>
                                    </div>
                                    <div className="d-flex justify-content-between border-top pt-1 mt-1 px-1">
                                        <span className="d-flex align-items-center">
                                            <i className="fas fa-calendar-days me-2 text-secondary" style={{width: "14px", textAlign: "center"}}></i>
                                            {t.start === "TBD" || t.start === "" ? "TBD" : t.start}
                                        </span>
                                        {t.start !== "TBD" && t.start !== "" && (
                                            <>
                                                <span className="mx-1 text-muted"><i className="fas fa-arrow-right" style={{fontSize: "0.7rem"}}></i></span>
                                                <span>{t.end}</span>
                                            </>
                                        )}
                                    </div>
                                </div>
                            </Card.Body>
                        </Card>
                    ))}
                </div>
            )}

            {/* MODALS */}
            <Modal show={showForm} onHide={() => setShowForm(false)} centered>
                <Modal.Header closeButton className="py-2 bg-light">
                    <Modal.Title style={{fontSize: "1rem"}}>
                        {formMode === "edit" ? "Edit Task" : formMode === "sub" ? "New Sub-Task" : "New Task"}
                    </Modal.Title>
                </Modal.Header>
                <Modal.Body className="p-3">
                    {formMode === "sub" && parentTask && (
                        <Alert variant="info" className="py-1 px-2 small mb-3">
                            Adding under: <strong>{parentTask.name}</strong> ({parentTask.id})
                        </Alert>
                    )}
                    <Form.Group className="mb-2">
                        <Form.Label className="small fw-bold text-muted">TASK NAME</Form.Label>
                        <Form.Control size="sm" type="text" value={formData.name} onChange={(e) => setFormData({...formData, name: e.target.value})} />
                    </Form.Group>
                    
                    <Form.Group className="mb-2">
                        <Form.Label className="small fw-bold text-muted">ASSIGN LEAD</Form.Label>
                        <Form.Select size="sm" value={formData.lead} onChange={(e) => setFormData({...formData, lead: e.target.value})}>
                            <option value="">Unassigned</option>
                            {/* Use First Name for Value, Full Name for Display */}
                            {teamMembers.map((m, i) => <option key={i} value={m.first}>{m.full}</option>)}
                        </Form.Select>
                    </Form.Group>

                    <Row className="g-2 mb-2">
                        <Col xs={6}>
                            <Form.Group>
                                <Form.Label className="small fw-bold text-muted">START</Form.Label>
                                <Form.Control size="sm" type="date" value={formData.start} onChange={(e) => setFormData({...formData, start: e.target.value})} />
                            </Form.Group>
                        </Col>
                        <Col xs={6}>
                            <Form.Group>
                                <Form.Label className="small fw-bold text-muted">END (EST)</Form.Label>
                                <Form.Control size="sm" type="date" value={formData.end} onChange={(e) => setFormData({...formData, end: e.target.value})} />
                            </Form.Group>
                        </Col>
                    </Row>

                    {formMode === "edit" && (
                        <Form.Group className="mb-2">
                            <Form.Label className="small fw-bold text-muted">% COMPLETE</Form.Label>
                            <Form.Select size="sm" value={formData.percent} onChange={(e) => setFormData({...formData, percent: e.target.value})}>
                                <option value="0%">0%</option>
                                <option value="25%">25%</option>
                                <option value="50%">50%</option>
                                <option value="75%">75%</option>
                                <option value="100%">100%</option>
                            </Form.Select>
                        </Form.Group>
                    )}

                    {status && <div className="text-danger small mt-2">{status}</div>}
                </Modal.Body>
                <Modal.Footer className="py-1 bg-light">
                    <Button variant="secondary" size="sm" onClick={() => setShowForm(false)}>Cancel</Button>
                    <Button variant="primary" size="sm" onClick={handleSave}>Save Task</Button>
                </Modal.Footer>
            </Modal>

            <Modal show={showDelete} onHide={() => setShowDelete(false)} centered size="sm">
                <Modal.Header closeButton className="py-2 bg-light">
                    <Modal.Title style={{fontSize: "1rem"}} className="fw-bold text-danger">Delete Task?</Modal.Title>
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

window.ProjectTasks = ProjectTasks;
