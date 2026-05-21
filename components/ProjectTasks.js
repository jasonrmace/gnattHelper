/* global React, ReactBootstrap, Excel */

const { useState, useEffect } = React;
const { Button, Card, Badge, Spinner, Modal, ButtonGroup } = window.ReactBootstrap || {};

const ProjectTasks = ({ project, onBack }) => {
    const [tasks, setTasks] = useState([]);
    const [isLoading, setIsLoading] = useState(false);
    
    // Modal States
    const [showDelete, setShowDelete] = useState(false);
    const [taskToDelete, setTaskToDelete] = useState(null);
    const [showFeatureModal, setShowFeatureModal] = useState(false); // For "Add/Edit" placeholder

    // --- 1. FETCH TASKS ---
    const fetchTasks = async () => {
        setIsLoading(true);
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");
                
                // Find Footer
                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                footerRange.load("rowIndex");
                await context.sync();

                // Read all data (Row 8 onwards)
                const dataStartIndex = 7;
                const rowCount = footerRange.rowIndex - dataStartIndex;
                
                if (rowCount <= 0) {
                    setTasks([]);
                    return;
                }

                const range = sheet.getRangeByIndexes(dataStartIndex, 0, rowCount, 8);
                range.load("text");
                await context.sync();

                // Filter for THIS project's children
                const rawRows = range.text;
                const projectTasks = [];
                const parentIdPrefix = `${project.id}.`; // e.g., "1."
                
                rawRows.forEach((row, index) => {
                    const idStr = row[0].toString();
                    
                    // Check if it belongs to this project (starts with "1.") 
                    // AND is not the project itself (not "1")
                    // AND is not a different project (not "10" or "11")
                    if (idStr.startsWith(parentIdPrefix) && idStr !== project.id.toString()) {
                        
                        // Calculate Depth (1.1 -> 0 indent, 1.1.1 -> 1 indent)
                        // We count the dots. "1.1" has 1 dot. "1.1.1" has 2 dots.
                        const dotCount = (idStr.match(/\./g) || []).length;
                        const depth = Math.max(0, dotCount - 1);

                        projectTasks.push({
                            id: idStr,
                            rowIndex: dataStartIndex + index,
                            name: row[1],
                            lead: row[2],
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

    // --- 2. DELETE TASK ---
    const handleDelete = async () => {
        if (!taskToDelete) return;
        
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");
                const range = sheet.getRangeByIndexes(taskToDelete.rowIndex, 0, 1, 1).getEntireRow();
                range.delete(Excel.DeleteShiftDirection.up);
                await context.sync();
                
                setShowDelete(false);
                setTaskToDelete(null);
                fetchTasks(); // Refresh list after delete
            });
        } catch (error) {
            console.error(error);
        }
    };

    // --- 3. EXCEL ACTIONS ---
    const handleJump = async (rowIndex) => {
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");
                sheet.activate();
                const range = sheet.getRangeByIndexes(rowIndex, 0, 1, 1).getEntireRow();
                range.select();
                await context.sync();
            });
        } catch (error) { console.error(error); }
    };

    return (
        <div className="mt-4">
            {/* HEADER */}
            <div className="d-flex align-items-center justify-content-between mb-3 pb-2 border-bottom">
                <div className="d-flex align-items-center">
                    <Button variant="light" size="sm" className="me-2 text-muted" onClick={onBack} title="Back to Projects">
                        <i className="fas fa-arrow-left"></i>
                    </Button>
                    <div style={{lineHeight: "1.1"}}>
                        <h6 className="m-0 fw-bold text-primary">Project {project.id}</h6>
                        <small className="text-muted" style={{fontSize: "0.7rem"}}>Task Manager</small>
                    </div>
                </div>
                <Button variant="primary" size="sm" onClick={() => setShowFeatureModal(true)}>
                    <i className="fas fa-plus me-1"></i> Add Task
                </Button>
            </div>

            {/* TASK LIST */}
            {isLoading ? (
                <div className="text-center py-5"><Spinner animation="border" size="sm" variant="primary" /></div>
            ) : tasks.length === 0 ? (
                <div className="text-center text-muted small mt-5">
                    <i className="fas fa-clipboard-list fa-2x mb-2 text-secondary opacity-50"></i><br/>
                    No tasks found for this project.
                </div>
            ) : (
                <div style={{ maxHeight: "calc(100vh - 220px)", overflowY: "auto", paddingRight: "4px" }}>
                    {tasks.map((t, idx) => (
                        <Card key={idx} className="mb-2 shadow-sm border-0" 
                              style={{ 
                                  marginLeft: `${t.depth * 24}px`, 
                                  borderLeft: t.depth > 0 ? "3px solid #e9ecef" : "none" 
                              }}>
                            <Card.Body className="p-2">
                                <div className="d-flex justify-content-between align-items-start">
                                    <div className="d-flex align-items-center" style={{overflow: "hidden"}}>
                                        <Badge bg="secondary" className="me-2" style={{fontSize: "0.65rem", minWidth: "35px"}}>{t.id}</Badge>
                                        <span className="fw-bold text-dark text-truncate" title={t.name} style={{fontSize: "0.85rem"}}>
                                            {t.name}
                                        </span>
                                    </div>
                                    
                                    {/* ACTION BUTTONS */}
                                    <ButtonGroup size="sm" className="ms-2 flex-shrink-0">
                                        <Button variant="light" className="px-2 text-primary" onClick={() => handleJump(t.rowIndex)} title="Locate in Excel">
                                            <i className="fas fa-location-arrow" style={{fontSize: "0.7rem"}}></i>
                                        </Button>
                                        <Button variant="light" className="px-2 text-secondary" onClick={() => setShowFeatureModal(true)} title="Edit Task">
                                            <i className="fas fa-pencil" style={{fontSize: "0.7rem"}}></i>
                                        </Button>
                                        <Button variant="light" className="px-2 text-danger" onClick={() => { setTaskToDelete(t); setShowDelete(true); }} title="Delete Task">
                                            <i className="fas fa-trash" style={{fontSize: "0.7rem"}}></i>
                                        </Button>
                                    </ButtonGroup>
                                </div>

                                {/* META DATA */}
                                <div className="mt-2 small text-muted">
                                    <div className="d-flex justify-content-between mb-1 align-items-center">
                                        <span className="d-flex align-items-center">
                                            <i className="fas fa-user me-2 text-secondary opacity-50" style={{width: "14px"}}></i> 
                                            {t.lead || "Unassigned"}
                                        </span>
                                        <Badge bg={t.percent === "100%" ? "success" : "light"} text={t.percent === "100%" ? "white" : "dark"} className="border fw-normal">
                                            {t.percent || "0%"}
                                        </Badge>
                                    </div>
                                    
                                    <div className="d-flex justify-content-between border-top pt-1 mt-1">
                                        <span className="d-flex align-items-center">
                                            <i className="fas fa-calendar-days me-2 text-secondary opacity-50" style={{width: "14px"}}></i> 
                                            {t.start || "TBD"}
                                        </span>
                                        {t.end && <span>➔ {t.end}</span>}
                                    </div>
                                </div>
                            </Card.Body>
                        </Card>
                    ))}
                </div>
            )}

            {/* DELETE CONFIRMATION MODAL */}
            <Modal show={showDelete} onHide={() => setShowDelete(false)} centered size="sm">
                <Modal.Header closeButton className="py-2 bg-light">
                    <Modal.Title style={{fontSize: "1rem"}} className="fw-bold text-danger">Delete Task?</Modal.Title>
                </Modal.Header>
                <Modal.Body className="small text-center py-4">
                    Are you sure you want to delete <br/>
                    <strong className="text-dark">{taskToDelete?.name}</strong>?
                    <div className="text-muted mt-2" style={{fontSize: "0.8em"}}>Row will be removed from Excel.</div>
                </Modal.Body>
                <Modal.Footer className="py-1 bg-light">
                    <Button variant="secondary" size="sm" onClick={() => setShowDelete(false)}>Cancel</Button>
                    <Button variant="danger" size="sm" onClick={handleDelete}>Yes, Delete</Button>
                </Modal.Footer>
            </Modal>

            {/* FEATURE COMING SOON MODAL (Replaces Alerts) */}
            <Modal show={showFeatureModal} onHide={() => setShowFeatureModal(false)} centered size="sm">
                <Modal.Body className="text-center py-4">
                    <i className="fas fa-tools fa-2x text-warning mb-3"></i>
                    <h6 className="fw-bold">Coming Soon</h6>
                    <p className="small text-muted mb-0">
                        The Edit/Add Task forms are being built. For now, please edit the row directly in Excel.
                    </p>
                </Modal.Body>
                <Modal.Footer className="py-1 justify-content-center">
                    <Button variant="primary" size="sm" onClick={() => setShowFeatureModal(false)}>Okay</Button>
                </Modal.Footer>
            </Modal>

        </div>
    );
};

window.ProjectTasks = ProjectTasks;
