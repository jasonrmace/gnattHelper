import React, { useState, useEffect } from 'react';
import { Container, Card, Form, Button, Alert } from 'react-bootstrap';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faUserCog, faUserCircle } from '@fortawesome/free-solid-svg-icons';
import { IdentityLogic } from '../utils/identityLogic';

const GANTT_MAIN_FILE = "Houston Summer 2026.xlsx";

const SettingsPage = ({ currentFileName }) => {
    const [currentIdentity, setCurrentIdentity] = useState('');
    const [teamMembers, setTeamMembers] = useState([]);
    const [selectedIdentity, setSelectedIdentity] = useState('');
    const [showSuccess, setShowSuccess] = useState(false);

    const canChangeIdentity = currentFileName === GANTT_MAIN_FILE;

    useEffect(() => {
        const fetchSettings = async () => {
            const identity = IdentityLogic.getIdentity();
            setCurrentIdentity(identity || 'Not set');
            setSelectedIdentity(identity || '');

            const members = await IdentityLogic.fetchTeamMembers();
            setTeamMembers(members);
        };
        fetchSettings();
    }, []);

    const handleIdentityChange = (e) => {
        setSelectedIdentity(e.target.value);
    };

    const handleSaveIdentity = () => {
        if (selectedIdentity) {
            IdentityLogic.setIdentity(selectedIdentity);
            setCurrentIdentity(selectedIdentity);
            setShowSuccess(true);
            window.GlobalToast.success(`Identity updated to ${selectedIdentity}!`);
            setTimeout(() => setShowSuccess(false), 3000);
        }
    };

    return (
        <Container className="mt-4">
            <Card className="shadow-sm">
                <Card.Body>
                    <h5 className="text-primary fw-bold mb-3"><FontAwesomeIcon icon={faUserCog} className="me-2" />User Settings</h5>

                    <Form.Group className="mb-4">
                        <Form.Label className="small fw-bold text-muted">WHO ARE YOU?</Form.Label>
                        <Form.Text className="text-muted d-block mb-2">
                                Identifying yourself allows the tool to highlight your specific tasks and handle your permissions.
                            </Form.Text>
                        
                        <Form.Select
                            size="sm"
                            value={selectedIdentity || ""}
                            onChange={handleIdentityChange}
                            disabled={!canChangeIdentity}
                        >
                            {!selectedIdentity && <option value="">Select your name...</option>}
                            
                            {/* Show the stored identity even if the Team list hasn't loaded or doesn't exist in this file */}
                            {selectedIdentity && !teamMembers.some(m => m.first === selectedIdentity) && (
                                <option value={selectedIdentity}>{selectedIdentity}</option>
                            )}

                            {teamMembers.map((member, idx) => (
                                <option key={idx} value={member.first}>
                                    {member.full}
                                </option>
                            ))}
                        </Form.Select>
                        {!canChangeIdentity && (
                            <Alert variant="info" className="mt-2 small">
                                You must load the file "Houston Summer 2026.xlsx" in order to change your identity. Ability to change in this file, coming soon!
                            </Alert>
                        )}
                        <Button variant="primary" size="sm" className="mt-2" onClick={handleSaveIdentity} disabled={!canChangeIdentity || !selectedIdentity}>Save Identity</Button>
                        {showSuccess && (
                            <div className="text-success small mt-2 fw-bold text-center">
                                Identity Saved!
                            </div>
                        )}
                    </Form.Group>
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