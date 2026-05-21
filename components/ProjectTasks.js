/* global React, ReactBootstrap, Excel */

const { useState, useEffect } = React;
const { Button, Card, Badge, Spinner, Modal, ButtonGroup } = window.ReactBootstrap || {};

const ProjectTasks = ({ project, onBack }) => {
    const [tasks, setTasks] = useState([]);
    const [isLoading, setIsLoading] = useState(false);
    const [showDelete, setShowDelete] = useState(false);
    const [taskToDelete, setTaskToDelete] = useState(null);

    // --- 1. FETCH TASKS ---
    const fetchTasks = async () => {
        setIsLoading(true);
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");
                
                // Find Footer to limit search
                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                footerRange.load("rowIndex");
                await context.sync();

                // Read all data (Row 8 onwards)
                const dataStartIndex = 7;
                const rowCount = footerRange.rowIndex - dataStartIndex;
                const range = sheet.getRangeByIndexes(dataStartIndex, 0, rowCount, 8);
                range.load("text");
                await context.sync();

                // Filter for THIS project's children
                // Logic: If Project ID is "1", we want "1.1", "1.2", "1.1.2", etc.
                const rawRows = range.text;
                const projectTasks = [];
                
                rawRows.forEach((row, index) => {
                    const idStr = row[0].toString();
                    const idVal = parseFloat(idStr);
                    
                    // Check if it belongs to this project (starts with "1." but isn't "1")
                    // And ensure it is not another integer project
                    if (idStr.startsWith(`${project.id}.`) && !Number.isInteger(idVal)) {
                        
                        // Calculate Depth (1.1 = Level 1, 1.1.1 = Level 2)
                        // We subtract 1 so 1.1 has depth 0 (no indent), 1.1.1 has depth 1
                        const depth = idStr.split('.').length - 2;

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
            console.error(error);
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
                fetchTasks(); // Refresh list
            });
        } catch (error) {
            console.error(error);
        }
    };

    // --- 3. JUMP TO ROW ---
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
                    <Button variant="light" size="sm" className="me-2 text-muted" onClick={onBack}>
                        <i className="fas fa-arrow-left"></i>
                    </Button>
                    <div>
                        <h6 className="m-0 fw-bold text-primary">Project {project.id}</h6>
                        <small className="text-muted">{project.name}</small>
                    </div>
                </div>
                <Button variant="primary" size="sm" onClick={() => alert("Add Task Logic Coming Soon")}>
                    <i className="fas fa-plus me-1"></i> Add Task
                </Button>
            </div>

            {/* TASK LIST */}
            {isLoading ? (
                <div className="text-center py-4"><Spinner animation="border" size="sm" /></div>
            ) : tasks.length === 0 ? (
                <div className="text-center text-muted small mt-4">No tasks found for this project.</div>
            ) : (
                <div style={{ maxHeight: "calc(100vh - 200px)", overflowY: "auto" }}>
                    {tasks.map((t, idx) => (
                        <Card key={idx} className="mb-2 shadow-sm border-0" style={{ marginLeft: `${t.depth * 20}px`, borderLeft: t.depth > 0 ? "3px solid #dee2e6" : "none" }}>
                            <Card.Body className="p-2">
                                <div className="d-flex justify-content-between align-items-start">
                                    <div className="d-flex align-items-center" style={{overflow: "hidden"}}>
                                        <Badge bg="secondary" className="me-2" style={{fontSize: "0.65rem"}}>{t.id}</Badge>
                                        <span className="fw-bold text-dark text-truncate" title={t.name} style={{fontSize: "0.9rem"}}>
                                            {t.name}
                                        </span>
                                    </div>
                                    
                                    {/* ACTIONS */}
                                    <ButtonGroup size="sm">
                                        <Button variant="light" className="px-2 text-primary" onClick={() => handleJump(t.rowIndex)} title="Locate">
                                            <i className="fas fa-location-arrow" style={{fontSize: "0.7rem"}}></i>
                                        </Button>
                                        <Button variant="light" className="px-2 text-danger" onClick={() => { setTaskToDelete(t); setShowDelete(true); }} title="Delete">
                                            <i className="fas fa-trash" style={{fontSize: "0.7rem"}}></i>
                                        </Button>
                                    </ButtonGroup>
                                </div>

                                <div className="mt-2 small text-muted">
                                    <div className="d-flex justify-content-between mb-1">
                                        <span><i className="fas fa-user me-2 text-secondary" style={{width: "14px"}}></i> {t.lead || "-"}</span>
                                        <Badge bg={t.percent === "100%" ? "success" : "light"} text={t.percent === "100%" ? "white" : "dark"} className="border">
                                            {t.percent || "0%"}
                                        </Badge>
                                    </div>
                                    
                                    <div className="d-flex justify-content-between border-top pt-1 mt-1">
                                        <span><i className="fas fa-calendar-days me-2 text-secondary"></i> {t.start || "TBD"}</span>
                                        {t.end && <span>➔ {t.end}</span>}
                                    </div>
                                </div>
                            </Card.Body>
                        </Card>
                    ))}
                </div>
            )}

            {/* DELETE MODAL */}
            <Modal show={showDelete} onHide={() => setShowDelete(false)} centered size="sm">
                <Modal.Header closeButton className="py-2">
                    <Modal.Title style={{fontSize: "1rem"}}>Delete Task?</Modal.Title>
                </Modal.Header>
                <Modal.Body className="small">
                    Are you sure you want to delete <strong>{taskToDelete?.name}</strong>? This cannot be undone.
                </Modal.Body>
                <Modal.Footer className="py-1">
                    <Button variant="secondary" size="sm" onClick={() => setShowDelete(false)}>Cancel</Button>
                    <Button variant="danger" size="sm" onClick={handleDelete}>Delete</Button>
                </Modal.Footer>
            </Modal>
        </div>
    );
};

window.ProjectTasks = ProjectTasks;
