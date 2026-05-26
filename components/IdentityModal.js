/* global React, ReactBootstrap */

const { useState, useEffect } = React;
const { Modal, Button, Form, Spinner } = ReactBootstrap;

const IdentityModal = ({ show, onHide }) => {
    const [teamList, setTeamList] = useState([]);
    const [selectedUser, setSelectedUser] = useState("");
    const [isLoading, setIsLoading] = useState(true);

    // Load Team Data when the Modal opens
    useEffect(() => {
        if (show && window.IdentityLogic) {
            const fetchData = async () => {
                setIsLoading(true);
                // Now returns objects: [{ first: "Chris", full: "Chris Smith" }]
                const members = await window.IdentityLogic.fetchTeamMembers();
                setTeamList(members);
                setIsLoading(false);
            };
            fetchData();
        }
    }, [show]);

    const handleSave = () => {
        if (window.IdentityLogic && selectedUser) {
            window.IdentityLogic.setIdentity(selectedUser);
            onHide(); // Close modal
        }
    };

    return (
        <Modal show={show} onHide={onHide} backdrop="static" keyboard={false} centered>
            <Modal.Header>
                <Modal.Title><i className="fas fa-user-tag me-2"></i>Welcome!</Modal.Title>
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

window.IdentityModal = IdentityModal;
