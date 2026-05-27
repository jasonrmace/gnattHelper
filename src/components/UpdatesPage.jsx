/* global React, Excel */
import React, { useState, useEffect } from 'react';
import { Container, Card, Badge, Spinner, Button } from 'react-bootstrap';
import toast from 'react-hot-toast';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faCodeCompare, faCheckDouble } from '@fortawesome/free-solid-svg-icons';
import { ChangelogLogic } from '../utils/changelogLogic';

const UpdatesPage = ({ onMarkedSeen }) => {
    const [logs, setLogs] = useState([]);
    const [isLoading, setIsLoading] = useState(true);

    const fetchLogs = async () => {
        await Excel.run(async (context) => {
            const data = await ChangelogLogic.getAllLogs(context);
            setLogs(data);
            setIsLoading(false);
        });
    };

    useEffect(() => {
        fetchLogs();
        
        // Mark all as seen when entering the page
        const markAll = async () => {
            await ChangelogLogic.markAllAsSeen();
            // Dismiss any active "nag" toasts (Admin updates or Summary) 
            // now that the user is viewing the full history.
            toast.dismiss();
            if (onMarkedSeen) onMarkedSeen();
        };
        markAll();
    }, []);

    if (isLoading) return (
        <div className="text-center py-5 mt-5">
            <Spinner animation="border" variant="primary" />
            <p className="text-muted mt-2">Loading updates...</p>
        </div>
    );

    return (
        <Container className="py-3">
            <div className="d-flex justify-content-between align-items-center mb-4">
                <h5 className="m-0 fw-bold text-dark">
                    <FontAwesomeIcon icon={faCodeCompare} className="text-primary me-2" />
                    Change History
                </h5>
            </div>

            {logs.length === 0 ? (
                <div className="text-center py-5 text-muted">No history found.</div>
            ) : (
                logs.map((log, index) => (
                    <Card key={index} className={`mb-2 border-0 shadow-sm ${log.isNew ? 'bg-light border-start border-primary' : ''}`} style={log.isNew ? {borderLeftWidth: '4px !important'} : {}}>
                        <Card.Body className="p-3 d-flex justify-content-between align-items-start">
                            <div className="flex-grow-1">
                                <p className="mb-1 text-dark" style={{fontSize: '0.9rem'}}>{log.text}</p>
                                <small className="text-muted" style={{fontSize: '0.75rem'}}>{log.timestamp}</small>
                            </div>
                            {log.isNew && <Badge bg="primary" className="ms-2">NEW</Badge>}
                        </Card.Body>
                    </Card>
                ))
            )}
        </Container>
    );
};

export default UpdatesPage;