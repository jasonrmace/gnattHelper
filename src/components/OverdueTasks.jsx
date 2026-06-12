/* global React, ReactBootstrap, Excel */
import React, { useState, useEffect } from 'react';
import { Container, Card, Badge, Spinner, Row, Col, ListGroup, Button } from 'react-bootstrap';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faClock, faUser, faTriangleExclamation, faLocationArrow } from '@fortawesome/free-solid-svg-icons';
import { IdentityLogic } from '../utils/identityLogic';

const OverdueTasks = ({ onNavigate }) => {
    const [tasks, setTasks] = useState([]);
    const [isLoading, setIsLoading] = useState(true);
    const currentUser = IdentityLogic.getIdentity();

    const fetchAllOverdueTasks = async () => {
        setIsLoading(true);
        try {
            await Excel.run(async (context) => {
                const sheets = ["Houston", "Dallas"];
                let allOverdue = [];
                const today = new Date();
                today.setHours(0, 0, 0, 0);

                for (const sheetName of sheets) {
                    const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                    const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                    footerRange.load(["rowIndex", "isNullObject"]);
                    sheet.load("isNullObject");
                    await context.sync();

                    if (sheet.isNullObject || footerRange.isNullObject) continue;

                    const dataStartIndex = 7;
                    const rowCount = footerRange.rowIndex - dataStartIndex;
                    if (rowCount <= 0) continue;

                    const range = sheet.getRangeByIndexes(dataStartIndex, 0, rowCount, 8);
                    range.load("text");
                    await context.sync();

                    const rawRows = range.text;
                    rawRows.forEach((row, index) => {
                        const idStr = row[0].toString();
                        const idNum = parseFloat(idStr);
                        
                        // Only process tasks (non-integers)
                        if (!isNaN(idNum) && !Number.isInteger(idNum)) {
                            const endDateStr = row[5];
                            const percent = row[7];
                            const taskEnd = new Date(endDateStr);

                            if (endDateStr !== "TBD" && !isNaN(taskEnd) && taskEnd < today && percent !== "100%") {
                                allOverdue.push({
                                    id: idStr,
                                    name: row[1].toString().replace(/^[↑\s]+/, ''),
                                    lead: row[2],
                                    end: endDateStr,
                                    percent: percent,
                                    location: sheetName,
                                    rowIndex: dataStartIndex + index,
                                    isMine: row[2] === currentUser
                                });
                            }
                        }
                    });
                }
                // Sort: User's tasks first, then by date
                allOverdue.sort((a, b) => {
                    if (a.isMine && !b.isMine) return -1;
                    if (!a.isMine && b.isMine) return 1;
                    return new Date(a.end) - new Date(b.end);
                });
                setTasks(allOverdue);
            });
        } catch (error) {
            console.error("Overdue fetch error:", error);
        } finally {
            setIsLoading(false);
        }
    };

    useEffect(() => {
        fetchAllOverdueTasks();
    }, []);

    const handleJump = async (task) => {
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(task.location);
                sheet.activate();
                const range = sheet.getRangeByIndexes(task.rowIndex, 0, 1, 1).getEntireRow();
                range.select();
                await context.sync();
            });
            if (onNavigate) onNavigate(task.location, task.id);
        } catch (error) { console.error(error); }
    };

    if (isLoading) return <div className="text-center py-5"><Spinner animation="border" variant="primary" /></div>;

    return (
        <Container className="py-3">
            <div className="d-flex justify-content-between align-items-center mb-4">
                <h5 className="m-0 fw-bold text-dark">
                    <FontAwesomeIcon icon={faClock} className="text-danger me-2" />
                    Overdue Project Tasks
                </h5>
                <Badge bg="danger" pill>{tasks.length} Total</Badge>
            </div>

            {tasks.length === 0 ? (
                <div className="text-center py-5 text-muted">No overdue tasks found! Great job.</div>
            ) : (
                <ListGroup variant="flush">
                    {tasks.map((t, idx) => (
                        <Card key={idx} className={`mb-2 border-0 shadow-sm ${t.isMine ? 'border-start border-primary border-4' : ''}`}>
                            <Card.Body className="p-3">
                                <Row className="align-items-center">
                                    <Col>
                                        <div className="d-flex align-items-center mb-1">
                                            <Badge bg="secondary" className="me-2 small" style={{fontSize: '0.65rem'}}>{t.id}</Badge>
                                            <span className={`fw-bold ${t.isMine ? 'text-primary' : ''}`} style={{fontSize: '0.9rem'}}>{t.name}</span>
                                            {t.isMine && <Badge bg="primary" className="ms-2" style={{fontSize: '0.6rem'}}>YOURS</Badge>}
                                        </div>
                                        <div className="d-flex gap-3 small text-muted mt-2" style={{fontSize: '0.75rem'}}>
                                            <span><FontAwesomeIcon icon={faUser} className="me-1 opacity-50" /> {t.lead || "Unassigned"}</span>
                                            <span className="text-danger fw-bold"><FontAwesomeIcon icon={faTriangleExclamation} className="me-1" /> Due: {t.end}</span>
                                            <span><Badge bg="light" text="dark" className="border">{t.location}</Badge></span>
                                            <span className="fw-bold text-dark">{t.percent} Complete</span>
                                        </div>
                                    </Col>
                                    <Col xs="auto">
                                        <Button 
                                            variant="light" 
                                            size="sm" 
                                            className="text-primary" 
                                            onClick={() => handleJump(t)}
                                            title="Go to Task"
                                        >
                                            <FontAwesomeIcon icon={faLocationArrow} />
                                        </Button>
                                    </Col>
                                </Row>
                            </Card.Body>
                        </Card>
                    ))}
                </ListGroup>
            )}
        </Container>
    );
};

export default OverdueTasks;