/* global React, ReactBootstrap, Excel */
import React, { useState, useEffect } from 'react';
import { Button, Form, Spinner, Row, Col, Alert } from 'react-bootstrap';

import { GanttLogic } from '../utils/ganttLogic';
import { ChangelogLogic } from '../utils/changelogLogic';
import { FormattingLogic } from '../utils/formattingLogic_v2';
import { VisualLogic } from '../utils/visualLogic';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faFolderPlus } from '@fortawesome/free-solid-svg-icons';

const CreateProject = ({ onProjectCreated }) => {
    // --- STATE ---
    const [teamMembers, setTeamMembers] = useState([]);
    const [formData, setFormData] = useState({ name: "", lead: "", startDate: "", endDate: "", projectNumber: "", isServiceCall: false });
    const [status, setStatus] = useState({ msg: "", variant: "light" });
    const [isLoading, setIsLoading] = useState(false);

    // --- 1. FETCH TEAM MEMBERS (As Objects) ---
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
                        const first = rows[i][0]?.trim();
                        const last = rows[i][1]?.trim();
                        if (first) {
                            members.push({ 
                                first: first,       // VALUE (for Excel)
                                full: `${first} ${last}` // LABEL (for UI)
                            });
                        }
                    }
                    setTeamMembers(members);
                });
            } catch (error) { console.error("Failed to load team:", error); }
        };
        fetchTeam();
    }, []);

    // --- 2. HANDLE INPUT CHANGES ---
    const handleChange = (field, value) => {
        setFormData(prev => ({ ...prev, [field]: value }));
    };

    // --- 3. CREATE LOGIC ---
    const handleCreate = async () => {
        if (!formData.name) {
            setStatus({ msg: "Project Name is required.", variant: "danger" });
            return;
        }
        if (!formData.isServiceCall && !formData.projectNumber) {
            setStatus({ msg: "Project Number is required.", variant: "danger" });
            return;
        }
        if (formData.startDate && formData.endDate && formData.startDate > formData.endDate) {
            setStatus({ msg: "End Date cannot be before Start Date.", variant: "danger" });
            return;
        }

        setIsLoading(true);
        setStatus({ msg: "Processing...", variant: "primary" });

        let createdId = null;
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("GanttChart");

                // A. PREPARE SHEET (Clear Filters - Matches VBA Safety)
                // Important: Inserting rows while filters are active can hide the new row
                const autoFilter = sheet.autoFilter;
                autoFilter.clearCriteria();
                await context.sync();

                // B. SAFE TEMPLATE SEARCH
                let namedItem = sheet.names.getItemOrNullObject("Level1Task");
                await context.sync();
                if (namedItem.isNullObject) {
                    namedItem = context.workbook.names.getItemOrNullObject("Level1Task");
                    await context.sync();
                }
                if (namedItem.isNullObject) throw new Error("Template 'Level1Task' not found.");
                
                const sourceRow = namedItem.getRange().getEntireRow();

                // C. FIND FOOTER & LOCATION
                const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { 
                    completeMatch: false, 
                    matchCase: false,
                    searchDirection: Excel.SearchDirection.forward
                });
                footerRange.load("rowIndex");
                await context.sync();
                
                if (footerRange.isNullObject) throw new Error("Critical: 'DO NOT DELETE' footer not found.");
                
                const insertRowIndex = footerRange.rowIndex;

                // D2. FIND PROJECT NUMBER COLUMN (Dynamic Lookup in Row 7)
                const headerRow = sheet.getRange("7:7");
                const projNumHeader = headerRow.find("Project Number", { completeMatch: true, matchCase: false });
                projNumHeader.load(["columnIndex", "isNullObject"]);
                await context.sync();

                let projNumCol = 3; // Fallback to index 3 (Col D) if header not found
                if (!projNumHeader.isNullObject) {
                    projNumCol = projNumHeader.columnIndex;
                }

                // D. CALCULATE NEXT ID (Matches VBA Logic)
                // Look at cell above footer. If numeric, add 1. Else 1.
                const lastIdCell = sheet.getCell(insertRowIndex - 1, 0);
                lastIdCell.load("values");
                await context.sync();

                let newId = 1;
                const lastVal = lastIdCell.values[0][0];
                if (!isNaN(lastVal) && lastVal !== "") {
                    newId = Math.floor(lastVal) + 1;
                }
                createdId = newId;

                // E. INSERT ROW & COPY TEMPLATE
                const insertRange = sheet.getRange(`${insertRowIndex + 1}:${insertRowIndex + 1}`);
                insertRange.insert(Excel.InsertShiftDirection.down);
                
                const newRow = sheet.getRange(`${insertRowIndex + 1}:${insertRowIndex + 1}`);
                newRow.copyFrom(sourceRow, Excel.RangeCopyType.all);

                // F. CALCULATE DURATION
                let duration = "";
                if (formData.startDate && formData.endDate) {
                    const start = new Date(formData.startDate);
                    const end = new Date(formData.endDate);
                    const diffTime = Math.abs(end - start);
                    // Added +1 for inclusive date calculation (consistent with ProjectTasks.js)
                    duration = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
                }

                // G. WRITE DATA
                // Col A (ID)
                sheet.getCell(insertRowIndex, 0).values = [[newId]];
                // Col B (Name)
                sheet.getCell(insertRowIndex, 1).values = [[formData.name]];
                
                // Project Number (Dynamic Column)
                if (formData.projectNumber) {
                    sheet.getCell(insertRowIndex, projNumCol).values = [[formData.projectNumber]];
                } else {
                    sheet.getCell(insertRowIndex, projNumCol).clear(Excel.ClearApplyTo.contents);
                }

                // Col C (Lead)
                if (formData.lead) {
                    sheet.getCell(insertRowIndex, 2).values = [[formData.lead]];
                } else {
                    sheet.getCell(insertRowIndex, 2).clear(Excel.ClearApplyTo.contents);
                }

                // Col E (Start)
                if (formData.startDate) {
                    sheet.getCell(insertRowIndex, 4).values = [[formData.startDate]];
                }

                // Col G (Duration)
                if (duration !== "") {
                    sheet.getCell(insertRowIndex, 6).values = [[duration]];
                }

                // H. TRIGGER LOGIC ENGINES (Phase 2 Integration)
                // Ensures internal formulas are correct immediately
                await GanttLogic.updateProjectAverages(context);

                // J. TARGETED FORMATTING & VISUALS
                const newGridRange = sheet.getRangeByIndexes(insertRowIndex, 10, 1, 365);
                const newNameRange = sheet.getRangeByIndexes(insertRowIndex, 2, 1, 1);
                await FormattingLogic.applyRulesToRange(context, newGridRange, newNameRange);
                await VisualLogic.refreshRange(context, insertRowIndex, 1);

                // Record the change
                await ChangelogLogic.logChange(context, `Created Project: ${formData.name} (ID: ${newId})`);

                // I. ACTIVATE & SELECT
                sheet.activate();
                newRow.select();
                await context.sync();

                setStatus({ msg: `Project "${formData.name}" (ID: ${newId}) Created!`, variant: "success" });
                setFormData({ name: "", lead: "", startDate: "", endDate: "", projectNumber: "", isServiceCall: false });
            });

            // Trigger refresh AFTER the Excel.run transaction is fully complete
            if (onProjectCreated) {
                onProjectCreated(createdId);
            }
        } catch (error) {
            console.error(error);
            setStatus({ msg: error.message, variant: "danger" });
        } finally {
            setIsLoading(false);
            setTimeout(() => setStatus({ msg: "", variant: "light" }), 4000);
        }
    };

    return (
        <div className="bg-light p-3 rounded mb-4 border" aria-label="New Project Form">
            <h6 className="fw-bold text-primary mb-3"><FontAwesomeIcon icon={faFolderPlus} className="me-2" />New Project</h6>

            {/* PROJECT NAME */}
            <Form.Group className="mb-2">
                <Form.Label className="small fw-bold text-muted d-flex justify-content-between">
                    PROJECT NAME
                </Form.Label>
                <Form.Control 
                    size="sm" 
                    type="text" 
                    placeholder="Enter name..." 
                    value={formData.name}
                    onChange={(e) => handleChange("name", e.target.value)}
                    disabled={isLoading}
                />
                <span className="text-danger ps-2" style={{fontSize: '0.65rem'}}>* REQUIRED</span>
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
                        <option key={idx} value={member.first}>
                            {member.full}
                        </option>
                    ))}
                </Form.Select>
            </Form.Group>

            {/* DATES */}
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

            {/* PROJECT NUMBER */}
            <Form.Group className="mb-2">
                <Form.Label className="small fw-bold text-muted d-flex justify-content-between">
                    PROJECT NUMBER
                </Form.Label>
                <Form.Control 
                    size="sm" 
                    type="text" 
                    placeholder="Enter project number..." 
                    value={formData.projectNumber}
                    onChange={(e) => handleChange("projectNumber", e.target.value)}
                    disabled={isLoading}
                />
                {!formData.isServiceCall && <span className="text-danger ps-2" style={{fontSize: '0.65rem'}}>* REQUIRED</span>}
                {formData.isServiceCall && <Form.Text className="text-muted ps-2" style={{fontSize: '0.65rem'}}>Optional for service calls</Form.Text>}
            </Form.Group>

            {/* SERVICE CALL TOGGLE */}
            <Form.Group className="mb-3">
                <Form.Check 
                    type="checkbox"
                    id="service-call-check"
                    label={<span className="small fw-bold">This is a Service Call</span>}
                    checked={formData.isServiceCall}
                    onChange={(e) => handleChange("isServiceCall", e.target.checked)}
                    disabled={isLoading}
                />
            </Form.Group>

            {/* SUBMIT BUTTON */}
            <Button 
                variant="primary" 
                size="sm" 
                className="w-100 shadow-sm" 
                onClick={handleCreate}
                disabled={isLoading || !formData.name || (!formData.isServiceCall && !formData.projectNumber)}
            >
                {isLoading ? (
                    <> {/* Loading spinner and text */}
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

export default CreateProject;
