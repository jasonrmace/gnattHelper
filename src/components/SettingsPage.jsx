/* global React, ReactBootstrap */
import React, { useState, useEffect }from 'react';
import { Container, Form, Button, Alert, Spinner, Card } from 'react-bootstrap';

import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faUserCog, faCheckCircle } from '@fortawesome/free-solid-svg-icons';
import { IdentityLogic } from '../utils/identityLogic';

const SettingsPage = () => {
    const [teamList, setTeamList] = useState([]);
    const [selectedUser, setSelectedUser] = useState("");
    const [isLoading, setIsLoading] = useState(true);
    const [showSuccess, setShowSuccess] = useState(false);

    // 1. Load Data on Mount
    useEffect(() => {
        const loadData = async () => {
            // A. Fetch Team Members (Now returns { first, full })
            const members = await IdentityLogic.fetchTeamMembers();
            setTeamList(members);

            // B. Get Current Saved Identity
            const current = IdentityLogic.getIdentity();
            if (current) setSelectedUser(current);
            
            setIsLoading(false);
        };
        loadData();
    }, []);

    // 2. Handle Save
    const handleSave = () => {
        IdentityLogic.setIdentity(selectedUser);
        setShowSuccess(true);
        
        // Hide success message after 3 seconds
        setTimeout(() => setShowSuccess(false), 3000);
    };

    if (isLoading) {
        return (
            <Container className="text-center mt-5">
                <Spinner animation="border" variant="primary" />
                <p className="mt-2 text-muted">Loading Team Data...</p>
            </Container>
        );
    }

    return (
        <Container className="mt-4" aria-label="User Settings Page">
            <h4 className="mb-4"><FontAwesomeIcon icon={faUserCog} className="me-2" /> User Settings</h4>
            
            <Card className="shadow-sm">
                <Card.Body>
                    <Form>
                        <Form.Group className="mb-3">
                            <Form.Label className="fw-bold">Who are you?</Form.Label>
                            <Form.Text className="text-muted d-block mb-2">
                                Identifying yourself allows the tool to highlight your specific tasks and handle your permissions.
                            </Form.Text>
                            
                            <Form.Select 
                                value={selectedUser} 
                                onChange={(e) => setSelectedUser(e.target.value)}
                            >
                                <option value="">-- Select Your Name --</option>
                                {teamList.map((m, index) => (
                                    <option key={index} value={m.first}>{m.full}</option>
                                ))}
                            </Form.Select>
                        </Form.Group>

                        {showSuccess && (
                            <Alert variant="success" className="py-2">
                                <FontAwesomeIcon icon={faCheckCircle} className="me-2" /> Identity Saved!
                            </Alert>
                        )}

                        <div className="d-grid gap-2">
                            <Button variant="primary" onClick={handleSave} disabled={!selectedUser}>
                                Save Identity
                            </Button>
                        </div>
                    </Form>
                </Card.Body>
            </Card>

            <div className="mt-4 text-center">
                <small className="text-muted">
                    Note: This setting is saved locally to this computer. If you switch devices, you may need to set this again.
                </small>
            </div>
        </Container>
    );
};

export default SettingsPage;
