const { useState } = React;
const { Button, Form,  Spinner } = ReactBootstrap;

const CreateProject = () => {

    const [projectName, setProjectName] = useState("");
    const [status, setStatus] = useState({ msg: "Ready.", type: "text-muted" });
    const [isLoading, setIsLoading] = useState(false);

    // --- CORE LOGIC (The Safe Search + Copy Row) ---
   const handleCreate = async () => {
        if (!projectName) {
            setStatus({ msg: "Please enter a project name.", variant: "danger" });
            return;
        }

        setIsLoading(true);
        setStatus({ msg: "Processing...", variant: "primary" });

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");

                // Safe Search
                let namedItem = sheet.names.getItemOrNullObject("Level1Task");
                await context.sync();
                if (namedItem.isNullObject) {
                    namedItem = context.workbook.names.getItemOrNullObject("Level1Task");
                    await context.sync();
                }
                if (namedItem.isNullObject) throw new Error("Named Range 'Level1Task' not found.");

                const sourceRow = namedItem.getRange().getEntireRow();
                sourceRow.load("rowIndex");
                
                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                footerRange.load("rowIndex");
                await context.sync();
                
                const footerIndex = footerRange.rowIndex;

                // Insert & Copy
                sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`).insert(Excel.InsertShiftDirection.down);
                sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`).copyFrom(sourceRow, Excel.RangeCopyType.all);

                // Update Name
                sheet.getCell(footerIndex, 1).values = [[projectName]];

                await context.sync();
                
                setStatus({ 
                    msg: `Success! Created '${projectName}'.`, 
                    variant: "success" 
                });
                setProjectName("");
            });
        } catch (error) {
            console.error(error);
            setStatus({ msg: error.message, variant: "danger" });
        } finally {
            setIsLoading(false);
        }
    };

    return(
        <>
            <div className="bg-light p-3 rounded mb-4 border">
                <Form.Group className="mb-2">
                    <Form.Label className="fw-bold small text-uppercase text-muted">Create New Project</Form.Label>
                    <Form.Control 
                        size="sm"
                        type="text" 
                        placeholder="Project Name..." 
                        value={projectName}
                        onChange={(e) => setProjectName(e.target.value)}
                        disabled={isLoading}
                    />
                </Form.Group>
                <Button 
                    variant="primary" 
                    size="sm" 
                    className="w-100" 
                    onClick={handleCreate} 
                    disabled={isLoading || !projectName}
                >
                    {isLoading ? (
                        <>
                            <Spinner as="span" animation="border" size="sm" role="status" aria-hidden="true" className="me-2" />
                            Working...
                        </>
                    ) : "Add Project"}
                </Button>
                {status.msg && <div className={`text-${status.variant} small mt-2 fw-bold text-center`}>{status.msg}</div>}
            </div>
        </>
    )
}

window.CreateProject = CreateProject;
    
