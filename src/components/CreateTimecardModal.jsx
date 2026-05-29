import React, { useState, useEffect } from 'react';
import { Modal, Button, Alert, Form } from 'react-bootstrap';
import DatePicker from "react-datepicker";
import { TimecardLogic } from '../utils/timecardLogic';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faCalendarPlus } from '@fortawesome/free-solid-svg-icons';

const CreateTimecardModal = ({ show, onHide, onCreated }) => {
    const [selectedDate, setSelectedDate] = useState(null);
    const [anchorDate, setAnchorDate] = useState(null);
    const [isProcessing, setIsProcessing] = useState(false);

    useEffect(() => {
        if (show) {
            const init = async () => {
                const latest = await TimecardLogic.getLatestPeriodEndDate();
                setAnchorDate(latest);
                const next = latest ? new Date(latest) : new Date();
                if (latest) next.setDate(next.getDate() + 14);
                else {
                    // Find next Friday if no sheets exist
                    while(next.getDay() !== 5) next.setDate(next.getDate() + 1);
                }
                setSelectedDate(next);
            };
            init();
        }
    }, [show]);

    const handleCreate = async () => {
        if (!selectedDate) return;
        setIsProcessing(true);
        try {
            const success = await TimecardLogic.createNewPeriod(selectedDate);
            if (success) {
                onHide();
                if (onCreated) onCreated();
            }
        } finally {
            setIsProcessing(false);
        }
    };

    // Calculate Start Date preview (Saturday)
    const getStartPreview = () => {
        if (!selectedDate) return "";
        const d = new Date(selectedDate);
        d.setDate(d.getDate() - 13);
        return d.toLocaleDateString();
    };

    /**
     * Filters the date picker to only allow every other Friday,
     * calculated from the latest existing timesheet in the workbook.
     */
    const filterBiweeklyFridays = (date) => {
        // Must be a Friday
        if (date.getDay() !== 5) return false;
        // If no sheets exist yet, allow any Friday to set the initial anchor
        if (!anchorDate) return true;

        const d1 = new Date(anchorDate.getFullYear(), anchorDate.getMonth(), anchorDate.getDate());
        const d2 = new Date(date.getFullYear(), date.getMonth(), date.getDate());
        const diffDays = Math.round(Math.abs(d2 - d1) / (1000 * 60 * 60 * 24));
        return diffDays % 14 === 0;
    };

    return (
        <Modal show={show} onHide={onHide} centered>
            <Modal.Header closeButton>
                <Modal.Title className="text-primary fw-bold" style={{fontSize: '1.1rem'}}>
                    <FontAwesomeIcon icon={faCalendarPlus} className="me-2" />
                    Add New Timesheet
                </Modal.Title>
            </Modal.Header>
            <Modal.Body>
                <p className="small text-muted mb-3">
                    Select the <strong>End Date (Friday)</strong> of the pay period you want to create. 
                    The tool will automatically copy the template and set the start date.
                </p>
                
                <Form.Group className="mb-3 text-center">
                    <Form.Label className="fw-bold small d-block mb-2">END DATE (FRIDAY)</Form.Label>
                    <DatePicker
                        selected={selectedDate}
                        onChange={date => setSelectedDate(date)}
                        filterDate={filterBiweeklyFridays}
                        className="form-control text-center"
                        dateFormat="MMMM d, yyyy"
                    />
                </Form.Group>

                <Alert variant="info" className="small py-2 mb-0">
                    <strong>Start Date:</strong> {getStartPreview()} (Saturday)
                </Alert>
            </Modal.Body>
            <Modal.Footer className="bg-light py-2">
                <Button variant="secondary" size="sm" onClick={onHide}>Cancel</Button>
                <Button variant="primary" size="sm" onClick={handleCreate} disabled={!selectedDate || isProcessing}>
                    {isProcessing ? "Creating..." : "Create Worksheet"}
                </Button>
            </Modal.Footer>
        </Modal>
    );
};

export default CreateTimecardModal;