/* global React, ReactBootstrap, Excel */

const { useState, useEffect } = React;
const { Container, Form, Button, Alert, Card, Badge, Stack, Spinner } = window.ReactBootstrap || {};

const ProjectList = () => {
    // --- STATE ---
    const [projectName, setProjectName] = useState("");
    const [projects, setProjects] = useState([]); 
    const [status, setStatus] = useState({ msg: "", variant: "light" });
    const [isCreating, setIsCreating] = useState(false);
    const [isFetching, setIsFetching] = useState(false);

    // --- 1. FETCH DATA (Updated for Row 8 Start) ---
    const fetchProjects = async () => {
        setIsFetching(true);
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");

                // A. Find the Bottom (Footer)
                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                footerRange.load("rowIndex");
                await context.sync();

                // B. Calculate Range 
                // Headers are Row 7 (Index 6). Data starts Row 8 (Index 7).
                const dataStartIndex = 7; 
                const footerIndex = footerRange.rowIndex;
                const rowCount = footerIndex - dataStartIndex;
                
                if (rowCount <= 0) {
                    setProjects([]); // No data rows exist between Header and Footer
                    return;
                }

                // C. Get Data (Columns A through H)
                // getRangeByIndexes(startRow, startCol, numRows, numCols)
                // Start at Index 7 (Row 8)
                const dataRange = sheet.getRangeByIndexes(dataStartIndex, 0, rowCount, 8);
                dataRange.load("text"); 
                await context.sync();

                // D. Map to Objects
                const formattedData = dataRange.text
                    .filter(row => row[1] !== "") // Skip empty names
                    .map((row) => ({
                        id: row[0],       // Col A
                        name: row[1],     // Col B
                        lead: row[2],     // Col C
                        start: row[4],    // Col E
                        end: row[5],      // Col F
                        percent: row[7]   // Col H
                    }));

                setProjects(formattedData);
            });
        } catch (error) {
            console.error(error);
        } finally {
            setIsFetching(false);
        }
    };

    // Load data on mount
    useEffect(() => {
        fetchProjects();
    }, []);

    // --- 2. CREATE PROJECT ---
    const handleCreate = async () => {
        if (!projectName) return;

        setIsCreating(true);
        setStatus({ msg: "Creating...", variant: "primary" });

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");
                
                let namedItem = sheet.names.getItemOrNullObject("Level1Task");
                await context.sync();
                if (namedItem.isNullObject) namedItem = context.workbook.names.getItemOrNullObject("Level1Task");
                if (namedItem.isNullObject) throw new Error("Template 'Level1Task' not found.");

                const sourceRow = namedItem.getRange().getEntireRow();
                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                footerRange.load("rowIndex");
                await context.sync();
                
                const footerIndex = footerRange.rowIndex;

                sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`).insert(Excel.InsertShiftDirection.down);
                sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`).copyFrom(sourceRow, Excel.RangeCopyType.all);
                sheet.getCell(footerIndex, 1).values = [[projectName]];

                await context.sync();
                
                setStatus({ msg: "Success!", variant: "success" });
                setProjectName("");
                
                // REFRESH LIST
                fetchProjects(); 
            });
        } catch (error) {
            setStatus({ msg: error.message, variant: "danger" });
        } finally {
            setIsCreating(false);
            setTimeout(() => setStatus({ msg: "", variant: "light" }), 3000);
        }
    };

    // --- 3. UI RENDER ---
    return (
        <>
            <div className="d-flex justify-content-between align-items-center mb-2">
                <h6 className="m-0 fw-bold text-primary">Active Projects ({projects.length})</h6>
                <Button variant="link" size="sm" className="text-decoration-none p-0" onClick={fetchProjects}>
                    {isFetching ? <Spinner animation="border" size="sm" /> : "Refresh"}
                </Button>
            </div>

            {projects.length === 0 && !isFetching && (
                <div className="text-center text-muted small mt-4">No projects found.</div>
            )}

            <div style={{ maxHeight: "400px", overflowY: "auto" }}>
                {projects.map((p, index) => (
                    <Card key={index} className="mb-2 shadow-sm border-0">
                        <Card.Body className="p-2">
                            <div className="d-flex justify-content-between align-items-start">
                                <div>
                                    <Badge bg="secondary" className="me-2">#{p.id}</Badge>
                                    <span className="fw-bold text-dark">{p.name}</span>
                                </div>
                                <Badge bg={p.percent === "100%" ? "success" : "info"} pill>
                                    {p.percent || "0%"}
                                </Badge>
                            </div>
                            
                            <div className="mt-2 small text-muted">
                                <div className="d-flex justify-content-between">
                                    <span>👤 {p.lead || "Unassigned"}</span>
                                </div>
                                <div className="d-flex justify-content-between border-top pt-1 mt-1">
                                    <span>📅 {p.start}</span>
                                    <span>➔ {p.end}</span>
                                </div>
                            </div>
                        </Card.Body>
                    </Card>
                ))}
            </div>
        </>
    );
};

window.ProjectList = ProjectList;
