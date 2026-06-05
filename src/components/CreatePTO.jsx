/* global React, ReactBootstrap, Excel */
import React, { useState, useEffect } from 'react';
import { Button, Form, Spinner, Row, Col, Alert } from 'react-bootstrap';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faCalendarPlus } from '@fortawesome/free-solid-svg-icons';

import { FormattingLogic } from '../utils/formattingLogic_v2';
import { VisualLogic } from '../utils/visualLogic';
import { ChangelogLogic } from '../utils/changelogLogic';
import { IdentityLogic } from '../utils/identityLogic';

const CreatePTO = ({ onPTOCreated }) => {
    const [teamMembers, setTeamMembers] = useState([]);
    const [formData, setFormData] = useState({ who: "", startDate: "", endDate: "" });
    const [status, setStatus] = useState({ msg: "", variant: "light" });
    const [isLoading, setIsLoading] = useState(false);

    const adminUsers = ["Rob", "Kevin", "Rob Kreps", "Kevin Rittner"];
    const isAdmin = adminUsers.includes(IdentityLogic.getIdentity());

    // 1. Fetch Team Members for dropdown and pre-select current user
    useEffect(() => {
        const fetchTeam = async () => {
            try {
                const members = await IdentityLogic.fetchTeamMembers();
                setTeamMembers(members);
                
                const currentIdentity = IdentityLogic.getIdentity();
                if (currentIdentity) {
                    setFormData(prev => ({ ...prev, who: currentIdentity }));
                }
            } catch (error) { console.error("Failed to load team:", error); }
        };
        fetchTeam();
    }, []);

    const handleChange = (field, value) => {
        setFormData(prev => ({ ...prev, [field]: value }));
    };

    const handleCreate = async () => {
        if (!formData.who || !formData.startDate || !formData.endDate) {
            setStatus({ msg: "All fields are required.", variant: "danger" });
            return;
        }
        if (formData.startDate > formData.endDate) {
            setStatus({ msg: "End Date cannot be before Start Date.", variant: "danger" });
            return;
        }

        setIsLoading(true);
        setStatus({ msg: "Adding to schedule...", variant: "primary" });

        try {
            await Excel.run(async (context) => {
                const table = context.workbook.tables.getItemOrNullObject("Vacations");
                table.load("isNullObject");
                await context.sync();

                if (table.isNullObject) {
                    throw new Error("The 'Vacations' table was not found in this workbook.");
                }

                // Calculate Duration (Total Days)
                const start = new Date(formData.startDate);
                const end = new Date(formData.endDate);
                const diffTime = Math.abs(end - start);
                const duration = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;

                // Add Row to Table
                // Columns: 1. Who, 2. Start Day, 3. Total Days, 4. End Day (Formula)
                // We pass null/empty string for column 4 so the table formula populates automatically
                table.rows.add(null, [
                    [formData.who, formData.startDate, duration, ""]
                ]);

                // Record the change
                await ChangelogLogic.logChange(context, `Added PTO for ${formData.who}: ${formData.startDate} to ${formData.endDate}`);

                await context.sync();

                // Trigger global sync for Houston and Dallas
                for (const name of ["Houston", "Dallas"]) {
                    await FormattingLogic.generateSmartRules(context, name);
                    await VisualLogic.refreshGridAlerts(context, name);
                }

                setStatus({ msg: "PTO Added Successfully!", variant: "success" });
                setFormData({ ...formData, startDate: "", endDate: "" });
                
                if (onPTOCreated) onPTOCreated();
            });
        } catch (error) {
            console.error(error);
            setStatus({ msg: error.message, variant: "danger" });
        } finally {
            setIsLoading(false);
            setTimeout(() => setStatus({ msg: "", variant: "light" }), 5000);
        }
    };

    return (
        <div className="bg-light p-3 rounded mb-4 border shadow-sm">
            <h6 className="fw-bold text-primary mb-3">
                <FontAwesomeIcon icon={faCalendarPlus} className="me-2" />
                Add PTO / Vacation
            </h6>

            {isAdmin && (
                <Form.Group className="mb-3">
                    <Form.Label className="small fw-bold text-muted">TEAM MEMBER</Form.Label>
                    <Form.Select 
                        size="sm" 
                        value={formData.who} 
                        onChange={(e) => handleChange("who", e.target.value)}
                        disabled={isLoading}
                    >
                        <option value="">Select name...</option>
                        {teamMembers.map((member, idx) => (
                            <option key={idx} value={member.first}>
                                {member.full}
                            </option>
                        ))}
                    </Form.Select>
                </Form.Group>
            )}

            <Row className="g-2 mb-3">
                <Col xs={6}>
                    <Form.Group>
                        <Form.Label className="small fw-bold text-muted">START DATE</Form.Label>
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
                        <Form.Label className="small fw-bold text-muted">END DATE</Form.Label>
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

            <Button 
                variant="primary" 
                size="sm" 
                className="w-100 fw-bold" 
                onClick={handleCreate}
                disabled={isLoading || !formData.who || !formData.startDate || !formData.endDate}
            >
                {isLoading ? <Spinner animation="border" size="sm" /> : "Submit Vacation"}
            </Button>

            {status.msg && (
                <div className={`text-${status.variant} small mt-2 fw-bold text-center`}>
                    {status.msg}
                </div>
            )}

            <div className="mt-4 pt-3 border-top">
                <p className="small mb-0 text-center">
                    <strong>Note:</strong> Please ensure all PTO and vacation requests are submitted through the normal processes outlined in <strong>Goldmine</strong> before adding them here. This is for <strong><i>Approved</i></strong> PTO/Vacation only.
                </p>
            </div>
        </div>
    );
};

export default CreatePTO;