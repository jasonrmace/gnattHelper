const { useState } = React;
const { Button, Form, Alert, Spinner } = ReactBootstrap;

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
            <Form.Group className="mb-3">
                <Form.Label className="fw-bold small text-uppercase">New Project Name</Form.Label>
                <Form.Control 
                    type="text" 
                    placeholder="Enter name..." 
                    value={projectName}
                    onChange={(e) => setProjectName(e.target.value)}
                    disabled={isLoading}
                />
            </Form.Group>

            <div className="d-grid gap-2">
                <Button 
                    variant="primary" 
                    onClick={handleCreate} 
                    disabled={isLoading}
                >
                    {isLoading ? (
                        <>
                            <Spinner as="span" animation="border" size="sm" role="status" aria-hidden="true" className="me-2" />
                            Working...
                        </>
                    ) : "Create Project"}
                </Button>
            </div>

            {/* Alert Component for Status */}
            <Alert variant={status.variant} className="mt-3 text-center small py-2 fw-bold">
                {status.msg}
            </Alert>
        </>
    )
}

export default CreateProject;
    
