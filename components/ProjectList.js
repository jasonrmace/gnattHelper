/* global React, ReactBootstrap, Excel */

const { useState, useEffect } = React;
const { Button, Card, Badge, Spinner } = window.ReactBootstrap || {};

const ProjectList = ({ refreshTrigger }) => {
    // --- STATE ---
    const [projects, setProjects] = useState([]);
    const [isFetching, setIsFetching] = useState(false);

    // --- DATA FETCHING ---
    const fetchProjects = async () => {
        setIsFetching(true);
        try {
            await Excel.run(async (context) => {
                // 1. Fetch Team Mapping Table First
                const teamSheet = context.workbook.worksheets.getItem("Team");
                const teamRange = teamSheet.getUsedRange();
                teamRange.load("text");

                // 2. Find Gantt Chart Footer
                const sheet = context.workbook.worksheets.getItem("GanttChart");
                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                footerRange.load("rowIndex");
                await context.sync();

                // 3. Build Team Lookup Map
                const teamMap = {};
                const teamRows = teamRange.text;
                for (let i = 1; i < teamRows.length; i++) {
                    const firstName = teamRows[i]?.[0]?.trim() || "";
                    const lastName = teamRows[i]?.[1]?.trim() || "";
                    if (firstName) {
                        teamMap[firstName.toLowerCase()] = `${firstName} ${lastName}`.trim();
                    }
                }

                // 4. Calculate Gantt Range (Row 8 start)
                const dataStartIndex = 7;
                const footerIndex = footerRange.rowIndex;
                const rowCount = footerIndex - dataStartIndex;

                if (rowCount <= 0) {
                    setProjects([]);
                    return;
                }

                // 5. Get Project Data
                const dataRange = sheet.getRangeByIndexes(dataStartIndex, 0, rowCount, 8);
                dataRange.load("text");
                await context.sync();

                // 6. AGGREGATE TASKS (New Logic)
                const allRows = dataRange.text.filter(row => row[1] !== ""); // Skip blank names
                const projectsMap = new Map();

                // Pass A: Identify Projects (Integers)
                allRows.forEach(row => {
                    const id = parseFloat(row[0]);
                    if (!isNaN(id) && Number.isInteger(id)) {
                        // It is a Project
                        const rawLead = row[2]?.trim() || "";
                        const fullLeadName = teamMap[rawLead.toLowerCase()] || rawLead;

                        projectsMap.set(id, {
                            id: row[0],
                            name: row[1],
                            lead: fullLeadName,
                            start: row[4],
                            end: row[5],
                            percent: row[7],
                            // Initialize Counters
                            totalTasks: 0,
                            completedTasks: 0
                        });
                    }
                });

                // Pass B: Count Tasks (Decimals)
                allRows.forEach(row => {
                    const id = parseFloat(row[0]);
                    if (!isNaN(id) && !Number.isInteger(id)) {
                        // It is a Task (e.g. 1.1), find parent (e.g. 1)
                        const parentId = Math.floor(id);
                        if (projectsMap.has(parentId)) {
                            const project = projectsMap.get(parentId);
                            project.totalTasks++;
                            // Check if task is 100% complete
                            if (row[7].includes("100%")) {
                                project.completedTasks++;
                            }
                        }
                    }
                });

                // Convert Map back to Array for Rendering
                setProjects(Array.from(projectsMap.values()));
            });
        } catch (error) {
            console.error("Fetch error:", error);
        } finally {
            setIsFetching(false);
        }
    };

    // Auto-refresh when trigger changes
    useEffect(() => {
        fetchProjects();
    }, [refreshTrigger]);

    // --- UI RENDER ---
    return (
        <div className="mt-4">
            <div className="d-flex justify-content-between align-items-center mb-2">
                <h6 className="m-0 fw-bold text-primary">Active Projects ({projects.length})</h6>
                <Button variant="link" size="sm" className="text-decoration-none p-0" onClick={fetchProjects}>
                    {isFetching ? <Spinner animation="border" size="sm" /> : "Refresh"}
                </Button>
            </div>

            {projects.length === 0 && !isFetching && (
                <div className="text-center text-muted small mt-2">No projects found.</div>
            )}

            <div style={{ maxHeight: "calc(100vh - 180px)", overflowY: "auto" }}>
                {projects.map((p, index) => (
                    <Card key={index} className="mb-2 shadow-sm border-0">
                        <Card.Body className="p-2">
                            <div className="d-flex justify-content-between align-items-start">
                                <div>
                                    <Badge bg="primary" className="me-2">#{p.id}</Badge>
                                    <span className="fw-bold text-dark">{p.name}</span>
                                </div>
                                <Badge bg={p.percent === "100%" ? "success" : p.percent === "0%" ? "danger" : "warning"} pill>
                                    {p.percent || "0%"}
                                </Badge>
                            </div>
                            
                            <div className="mt-2 small text-muted">
                                
                                {/* TASK COUNT ROW */}
                                <div className="d-flex justify-content-between mb-1 text-dark">
                                    <span>☑️ Tasks: <strong>{p.completedTasks}/{p.totalTasks}</strong> Complete</span>
                                </div>

                                <div className="d-flex justify-content-between">
                                    <span>👤 {p.lead || "Unassigned"}</span>
                                </div>
                                
                                <div className="d-flex justify-content-between border-top pt-1 mt-1">
                                    {p.start === "TBD" && (
                                        <>
                                            <span>📅 {p.start}</span>
                                        </>
                                    )}
                                    {p.start === "" && <span>📅 TBD</span>}
                                    {p.start !== "TBD" && p.start !== "" && (
                                        <>
                                            <span>📅 {p.start}</span>
                                            <span> ➔ </span>
                                            <span>{p.end}</span>
                                        </>
                                    )}
                                </div>
                            </div>
                        </Card.Body>
                    </Card>
                ))}
            </div>
        </div>
    );
};

window.ProjectList = ProjectList;
