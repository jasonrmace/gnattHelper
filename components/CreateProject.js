/* global React, ReactBootstrap, Excel */

const { useState, useEffect } = React;
const { Button, Form, Spinner, Row, Col, Alert } = ReactBootstrap;

const CreateProject = () => {
    // --- STATE ---
    const [teamMembers, setTeamMembers] = useState([]);
    const [formData, setFormData] = useState({
        name: "",
        lead: "",
        startDate: "",
        endDate: ""
    });
    const [status, setStatus] = useState({ msg: "", variant: "light" });
    const [isLoading, setIsLoading] = useState(false);

    // --- 1. FETCH TEAM MEMBERS ON LOAD ---
    useEffect(() => {
        const fetchTeam = async () => {
            try {
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getItem("Team");
                    const range = sheet.getUsedRange();
                    range.load("text");
                    await context.sync();

                    const rows = range.text;
                    const members = [];
                    // Skip header (row 0). Col A=First, Col B=Last
                    for (let i = 1; i < rows.length; i++) {
                        const first = rows[i][0];
                        const last = rows[i][1];
                        if (first) members.push(`${first} ${last}`.trim());
                    }
                    setTeamMembers(members);
                });
            } catch (error) {
                console.error("Failed to load team:", error);
            }
        };
        fetchTeam();
    }, []);

    // --- 2. HANDLE INPUT CHANGES ---
    const handleChange = (field, value) => {
        setFormData(prev => ({ ...prev, [field]: value }));
    };

    // --- 3. CREATE LOGIC ---
    const handleCreate = async () => {
        // Validation
        if (!formData.name) {
            setStatus({ msg: "Project Name is required.", variant: "danger" });
            return;
        }
        if (formData.startDate && formData.endDate && formData.startDate > formData.endDate) {
            setStatus({ msg: "End Date cannot be before Start Date.", variant: "danger" });
            return;
        }

        setIsLoading(true);
        setStatus({ msg: "Processing...", variant: "primary" });

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");

                // A. Safe Search for Template
                let namedItem = sheet.names.getItemOrNullObject("Level1Task");
                await context.sync();
                if (namedItem.isNullObject) {
                    namedItem = context.workbook.names.getItemOrNullObject("Level1Task");
                    await context.sync();
                }
                if (namedItem.isNullObject) throw new Error("Template 'Level1Task' not found.");

                const sourceRow = namedItem.getRange().getEntireRow();

                // B. Find Footer & Insert Location
                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                footerRange.load("rowIndex");
                await context.sync();
                
                const footerIndex = footerRange.rowIndex;
                const insertRange = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);

                // C. Insert & Copy Template
                insertRange.insert(Excel.InsertShiftDirection.down);
                // Re-grab the new row location
                const newRow = sheet.getRange(`${footerIndex + 1}:${footerIndex + 1}`);
                newRow.copyFrom(sourceRow, Excel.RangeCopyType.all);

                // D. Calculate Duration (Days) for Column G
                let duration = "";
                if (formData.startDate && formData.endDate) {
                    const start = new Date(formData.startDate);
                    const end = new Date(formData.endDate);
                    const diffTime = Math.abs(end - start);
                    duration = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
                }

                // E. Write Data
                // Col B (1) = Name
                sheet.getCell(footerIndex, 1).values = [[formData.name]];
                
                // Col C (2) = Lead
                if (formData.lead) {
                    sheet.getCell(footerIndex, 2).values = [[formData.lead]];
                }

                // Col E (4) = Start Date
                if (formData.startDate) {
                    // Format date for Excel (YYYY-MM-DD usually works best)
                    sheet.getCell(footerIndex, 4).values = [[formData.startDate]];
                }

                // Col G (6) = Duration (Calculated from End Date)
                if (duration !== "") {
                    sheet.getCell(footerIndex, 6).values = [[duration]];
                }

                // F. Activate Sheet & Highlight Row
                sheet.activate();
                newRow.select();

                await context.sync();

                setStatus({ msg: "Success!", variant: "success" });
                
                // Reset Form
                setFormData({ name: "", lead: "", startDate: "", endDate: "" });
                
                // Trigger Global Refresh (if needed by other components)
                // window.dispatchEvent(new Event('projectCreated')); 
            });
        } catch (error) {
            console.error(error);
            setStatus({ msg: error.message, variant: "danger" });
        } finally {
            setIsLoading(false);
            setTimeout(() => setStatus({ msg: "", variant: "light" }), 3000);
        }
    };

    return (
        <div className="bg-light p-3 rounded mb-4 border">
            <h6 className="fw-bold text-primary mb-3"><i className="fas fa-folder-plus me-2"></i>New Project</h6>
            
            {/* PROJECT NAME */}
            <Form.Group className="mb-2">
                <Form.Label className="small fw-bold text-muted">PROJECT NAME</Form.Label>
                <Form.Control 
                    size="sm" 
                    type="text" 
                    placeholder="Enter name..." 
                    value={formData.name} 
                    onChange={(e) => handleChange("name", e.target.value)} 
                    disabled={isLoading} 
                />
            </Form.Group>

            {/* PROJECT LEAD */}
            <Form.Group className="mb-2">
                <Form.Label className="small fw-bold text-muted">PROJECT LEAD</Form.Label>
                <Form.Select 
                    size="sm"
                    value={formData.lead}
                    onChange={(e) => handleChange("lead", e.target.value)}
                    disabled={isLoading}
                >
                    <option value="">Select Lead...</option>
                    {teamMembers.map((member, idx) => (
                        <option key={idx} value={member}>{member}</option>
                    ))}
                </Form.Select>
            </Form.Group>

            {/* DATES ROW */}
            <Row className="g-2 mb-3">
                <Col xs={6}>
                    <Form.Group>
                        <Form.Label className="small fw-bold text-muted">START</Form.Label>
                        <Form.Control 
                            size="sm" 
                            type="date" 
                            value={formData.startDate} 
                            onChange={(e) => handleChange("startDate", e.target.value)} 
                            disabled={isLoading} 
                        />
                    </Form.Group>
                </Col>
                <Col xs={6}>
                    <Form.Group>
                        <Form.Label className="small fw-bold text-muted">END (EST)</Form.Label>
                        <Form.Control 
                            size="sm" 
                            type="date" 
                            value={formData.endDate} 
                            onChange={(e) => handleChange("endDate", e.target.value)} 
                            disabled={isLoading} 
                        />
                    </Form.Group>
                </Col>
            </Row>

            {/* SUBMIT BUTTON */}
            <Button 
                variant="primary" 
                size="sm" 
                className="w-100 shadow-sm" 
                onClick={handleCreate} 
                disabled={isLoading || !formData.name}
            >
                {isLoading ? (
                    <>
                        <Spinner as="span" animation="border" size="sm" role="status" aria-hidden="true" className="me-2" />
                        Working...
                    </>
                ) : "Add Project"}
            </Button>

            {/* STATUS MESSAGE */}
            {status.msg && (
                <div className={`text-${status.variant} small mt-2 fw-bold text-center`}>
                    {status.msg}
                </div>
            )}
        </div>
    );
};

window.CreateProject = CreateProject;
