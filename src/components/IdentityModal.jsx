/* global React, ReactBootstrap */
import React, { useState, useEffect } from 'react';
import { Modal, Button, Form, Spinner } from 'react-bootstrap';

import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faUserTag } from '@fortawesome/free-solid-svg-icons';
import { IdentityLogic } from '../utils/identityLogic';

const IdentityModal = ({ show, onHide }) => {
    const [teamList, setTeamList] = useState([]);
    const [selectedUser, setSelectedUser] = useState("");
    const [isLoading, setIsLoading] = useState(true);

    // Load Team Data when the Modal opens
    useEffect(() => {
        if (show) {
            const fetchData = async () => {
                setIsLoading(true);
                // Now returns objects: [{ first: "Chris", full: "Chris Smith" }]
                const members = await IdentityLogic.fetchTeamMembers();
                setTeamList(members);
                setIsLoading(false);
            };
            fetchData();
        }
    }, [show]);

    const handleSave = () => {
        if (selectedUser) {
            IdentityLogic.setIdentity(selectedUser);
            onHide(); // Close modal
        }
    };

    return (
        <Modal show={show} onHide={onHide} backdrop="static" keyboard={false} centered aria-label="Identity Selection Modal">
            <Modal.Header>
                <Modal.Title><FontAwesomeIcon icon={faUserTag} className="me-2" />Welcome!</Modal.Title>
            </Modal.Header>
            <Modal.Body>
                <p className="text-muted small">
                    To help you manage your tasks, please identify yourself from the team list.
                </p>

                {isLoading ? (
                    <div className="text-center py-3">
                        <Spinner animation="border" size="sm" variant="primary" />
                        <span className="ms-2">Loading Team...</span>
                    </div>
                ) : (
                    <Form.Group>
                        <Form.Label className="fw-bold">I am...</Form.Label>
                        <Form.Select 
                            value={selectedUser} 
                            onChange={(e) => setSelectedUser(e.target.value)}
                        >
                            <option value="">-- Select Your Name --</option>
                            {/* UPDATED MAPPING LOGIC */}
                            {teamList.map((m, index) => (
                                <option key={index} value={m.first}>{m.full}</option>
                            ))}
                        </Form.Select>
                    </Form.Group>
                )}
            </Modal.Body>
            <Modal.Footer>
                <Button variant="secondary" onClick={onHide}>
                    Skip for Now
                </Button>
                <Button variant="primary" onClick={handleSave} disabled={!selectedUser}>
                    Save & Continue
                </Button>
            </Modal.Footer>
        </Modal>
    );
};

export default IdentityModal;
