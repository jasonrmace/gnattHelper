/* global React, ReactBootstrap, Excel */
import React from 'react';
import { Container, Row, Col, Card, Button } from 'react-bootstrap';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faCity, faBuilding, faArrowRight } from '@fortawesome/free-solid-svg-icons';

const HomePage = ({ setActiveTab }) => {
    const handleNavigate = (tab, sheetName) => {
        // 1. Update UI
        setActiveTab(tab);

        // 2. Sync Excel Worksheet
        Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            sheet.activate();
            await context.sync();
        }).catch(err => console.warn(`Home: Could not activate sheet "${sheetName}".`, err));
    };

    return (
        <Container className="py-4 animate-enter">
            <div className="text-center mb-5">
                <h4 className="fw-bold text-primary">Welcome to the Barbizon Light of Texas</h4>
                <p className="text-muted">Project Management Assistant</p>
                <hr className="w-25 mx-auto opacity-25" />
            </div>

            <Row className="g-4">
                {/* HOUSTON CARD */}
                <Col xs={12}>
                    <Card 
                        className="shadow-sm border-0 h-100 hover-shadow transition-all" 
                        style={{ cursor: 'pointer', borderLeft: '5px solid #ff4500' }}
                        onClick={() => handleNavigate("HoustonList", "Houston")}
                    >
                        <Card.Body className="d-flex align-items-center p-4">
                            <div className="bg-danger bg-opacity-10 p-3 rounded-circle me-4 text-danger">
                                <FontAwesomeIcon icon={faCity} size="2x" style={{ width: '40px' }} />
                            </div>
                            <div className="flex-grow-1">
                                <h5 className="fw-bold mb-1">Houston Projects</h5>
                                <p className="text-muted small mb-0">View and manage Houston office schedules</p>
                            </div>
                            <div className="text-muted opacity-50">
                                <FontAwesomeIcon icon={faArrowRight} />
                            </div>
                        </Card.Body>
                    </Card>
                </Col>

                {/* DALLAS CARD */}
                <Col xs={12}>
                    <Card 
                        className="shadow-sm border-0 h-100 hover-shadow transition-all" 
                        style={{ cursor: 'pointer', borderLeft: '5px solid #0d6efd' }}
                        onClick={() => handleNavigate("DallasList", "Dallas")}
                    >
                        <Card.Body className="d-flex align-items-center p-4">
                            <div className="bg-primary bg-opacity-10 p-3 rounded-circle me-4 text-primary">
                                <FontAwesomeIcon icon={faBuilding} size="2x" style={{ width: '40px' }} />
                            </div>
                            <div className="flex-grow-1">
                                <h5 className="fw-bold mb-1">Dallas Projects</h5>
                                <p className="text-muted small mb-0">View and manage Dallas office schedules</p>
                            </div>
                            <div className="text-muted opacity-50">
                                <FontAwesomeIcon icon={faArrowRight} />
                            </div>
                        </Card.Body>
                    </Card>
                </Col>
            </Row>

            <div className="mt-5 text-center opacity-50">
                <small className="text-muted">Select an office to begin managing projects and tasks.</small>
            </div>
        </Container>
    );
};

export default HomePage;