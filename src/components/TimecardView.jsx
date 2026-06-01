/* global Excel */
import React, { useState, useEffect } from 'react';
import { Container, Card, Spinner, ListGroup, Button, Badge } from 'react-bootstrap';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { 
    faCalendar, 
    faChevronRight, 
    faExclamationCircle 
} from '@fortawesome/free-solid-svg-icons';
import { faCalendarCheck } from '@fortawesome/free-regular-svg-icons';
import TimecardDetail from './TimecardDetail';
import { TimecardLogic } from '../utils/timecardLogic';

const TimecardView = ({ currentFileName, refreshTrigger }) => {
    const [timesheets, setTimesheets] = useState([]);
    const [loading, setLoading] = useState(true);
    const [selectedPeriod, setSelectedPeriod] = useState(null);

    const fetchTimecardData = async () => {
        setLoading(true);
        try {
            await Excel.run(async (context) => {
                const sheets = context.workbook.worksheets;
                sheets.load("items/name, items/tabColor");
                await context.sync();

                const sheetsToLoad = [];
                // Access worksheets from left to right, excluding the last one (Time Off Tracker)
                const count = sheets.items.length;
                for (let i = 0; i < count - 1; i++) {
                    const sheet = sheets.items[i];
                    if (sheet.name === "Time Off Tracker") continue;

                    const startDateRange = sheet.getRange("C7");
                    const endDateRange = sheet.getRange("I19");

                    // Validation Ranges: Checksums and Daily Blocks
                    const w1Data = sheet.getRangeByIndexes(12, 2, 5, 7); // Rows 13-17, Cols C-I
                    const w2Data = sheet.getRangeByIndexes(24, 2, 5, 7); // Rows 25-29, Cols C-I
                    const k17 = sheet.getRange("K17");
                    const k29 = sheet.getRange("K29");
                    const k34 = sheet.getRange("K34");
                    const h38 = sheet.getRange("H38");

                    startDateRange.load(["text", "values"]);
                    endDateRange.load(["text", "values"]);
                    w1Data.load("values");
                    w2Data.load("values");
                    k17.load("values");
                    k29.load("values");
                    k34.load("values");
                    h38.load("values");

                    sheetsToLoad.push({
                        name: sheet.name,
                        tabColor: sheet.tabColor,
                        startRange: startDateRange,
                        endRange: endDateRange,
                        w1Data, w2Data, k17, k29, k34, h38
                    });
                }

                await context.sync();

                const results = sheetsToLoad.map(item => {
                    // Detect Green Variants (Processed) vs Orange (Pending Submission)
                    const hex = item.tabColor ? item.tabColor.toUpperCase() : "";
                    const isProcessed = ["#00B050", "#70AD47", "#008000", "#92D050", "#C6EFCE"].includes(hex);
                    const isPending = ["#FFC000", "#FFFF00", "#FFD700", "#FFA500"].includes(hex);
                    
                    const isLocked = isProcessed || isPending;
                    
                    const now = new Date();
                    const todaySerial = Math.floor(now.getTime() / (24 * 60 * 60 * 1000) + 25569);
                    
                    // Determine if today falls within this period
                    let isCurrent = false;
                    let dueStatus = null; // friday, monday_due, monday_past_due
                    let mondayDateStr = "";

                    const startVal = item.startRange.values[0][0];
                    const endVal = item.endRange.values[0][0];

                    if (typeof startVal === 'number' && typeof endVal === 'number') {
                        if (todaySerial >= startVal && todaySerial <= endVal) {
                            isCurrent = true;
                        }

                        // Deadline Logic
                        if (todaySerial === endVal) {
                            // Today is the last Friday
                            dueStatus = 'friday';
                            const mDate = new Date(Math.round((endVal + 3 - 25569) * 86400 * 1000));
                            const localM = new Date(mDate.getTime() + (mDate.getTimezoneOffset() * 60000));
                            mondayDateStr = `${localM.getMonth() + 1}/${localM.getDate()}/${localM.getFullYear()}`;
                        } else if (!isLocked && todaySerial === endVal + 3) {
                            // Today is the Monday following the end of the timesheet
                            const currentHour = now.getHours();
                            dueStatus = currentHour < 12 ? 'monday_due' : 'monday_past_due';
                        }
                    }

                    let hasError = false;
                    // 1. Checksums
                    if (TimecardLogic.isErrorValue(item.k17.values[0][0]) || 
                        TimecardLogic.isErrorValue(item.k29.values[0][0]) || 
                        TimecardLogic.isErrorValue(item.k34.values[0][0]) || 
                        TimecardLogic.isErrorValue(item.h38.values[0][0])) {
                        hasError = true;
                    }

                    // 2. Scan Daily Data (Week 1 & 2)
                    [item.w1Data.values, item.w2Data.values].forEach(weekData => {
                        for (let col = 0; col < 7; col++) {
                            const total = parseFloat(weekData[0][col]) || 0;
                            const off = parseFloat(weekData[1][col]) || 0;
                            const trav = parseFloat(weekData[2][col]) || 0;
                            const pto = parseFloat(weekData[3][col]) || 0;
                            const ptoType = weekData[4][col];

                            if (total < 0) hasError = true;
                            // Allocation Mismatch
                            if (total > 0 && TimecardLogic.isErrorValue((off + trav + pto) - total)) hasError = true;
                            // Missing PTO Type
                            if (pto > 0 && (!ptoType || ptoType === "")) hasError = true;
                        }
                    });

                    return {
                        name: item.name,
                        start: item.startRange.text[0][0] || "TBD",
                        end: item.endRange.text[0][0] || "TBD",
                        isSubmitted: isLocked,
                        isProcessed: isProcessed,
                        isPending: isPending,
                        isCurrent: isCurrent,
                        dueStatus: dueStatus,
                        mondayDateStr: mondayDateStr,
                        hasError: hasError,
                        tabColor: hex
                    };
                });

                setTimesheets(results);
            });
        } catch (error) {
            console.error("Error fetching timesheet list:", error);
        } finally {
            setLoading(false);
        }
    };

    const handleActivateSheet = async (sheetName) => {
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(sheetName);
                sheet.activate();
                await context.sync();
            });
        } catch (error) {
            console.error("Could not activate sheet:", error);
        }
    };

    const handleSelectPeriod = (ts) => {
        handleActivateSheet(ts.name);
        setSelectedPeriod(ts);
    };

    useEffect(() => {
        fetchTimecardData();
    }, [refreshTrigger]);

    if (selectedPeriod) {
        return (
            <TimecardDetail period={selectedPeriod} onBack={() => {
                setSelectedPeriod(null);
                fetchTimecardData(); // Refresh badges when returning
            }} />
        );
    }

    return (
        <Container className="mt-4">
            <Card className="shadow-sm">
                <Card.Body className="p-3">
                    <div className="d-flex justify-content-between align-items-center mb-3">
                        <h5 className="text-primary fw-bold mb-0">Available Timesheets</h5>
                        <Button variant="link" size="sm" onClick={fetchTimecardData} className="p-0 text-decoration-none">Refresh</Button>
                    </div>

                    {loading ? (
                        <div className="text-center py-5">
                            <Spinner animation="border" variant="primary" />
                            <p className="text-muted small mt-2">Reading workbook...</p>
                        </div>
                    ) : timesheets.length === 0 ? (
                        <div className="text-center py-4 text-muted small">No pay periods detected.</div>
                    ) : (
                        <ListGroup variant="flush">
                            {timesheets.map((ts, idx) => (
                                <ListGroup.Item 
                                    key={idx} 
                                    action
                                    className={`d-flex justify-content-between align-items-center px-0 py-3 ${ts.isSubmitted ? 'opacity-75' : ''}`}
                                    onClick={() => handleSelectPeriod(ts)}
                                    style={{ cursor: 'pointer' }}
                                >
                                    <div className="d-flex align-items-center">
                                        <div className={`rounded p-2 me-3 text-primary`} style={{ width: '40px', textAlign: 'center' }}>
                                            <FontAwesomeIcon icon={ts.isSubmitted ? faCalendarCheck : faCalendar} />
                                        </div>
                                        <div>
                                            <div className="fw-bold text-dark" style={{ fontSize: '0.9rem' }}>
                                                {ts.start} — {ts.end}
                                                {ts.isCurrent && <Badge bg="info" className="ms-2" style={{fontSize: '0.6rem'}}>CURRENT WEEK</Badge>}
                                                {ts.dueStatus === 'friday' && (
                                                    <Badge bg="warning" text="dark" className="ms-2" style={{fontSize: '0.6rem'}}>Timesheet Due by Monday, {ts.mondayDateStr}</Badge>
                                                )}
                                                {ts.dueStatus === 'monday_due' && (
                                                    <Badge bg="danger" className="ms-2" style={{fontSize: '0.6rem'}}>Timesheet is Due Today!</Badge>
                                                )}
                                                {ts.dueStatus === 'monday_past_due' && (
                                                    <Badge bg="danger" className="ms-2" style={{fontSize: '0.6rem'}}>Timesheet is Past Due! Please Submit ASAP!</Badge>
                                                )}
                                            </div>
                                            {/* <small className="text-muted">{ts.name}</small> */}
                                            {ts.isProcessed && <Badge bg="success" className="ms-2" style={{fontSize: '0.6rem'}}>SUBMITTED & PROCESSED</Badge>}
                                            {ts.isPending && <Badge bg="warning" text="dark" className="ms-2" style={{fontSize: '0.6rem'}}>PENDING PROCESSING</Badge>}
                                            {ts.hasError && <Badge bg="danger" className="ms-2" style={{fontSize: '0.6rem'}}><FontAwesomeIcon icon={faExclamationCircle} className="me-1" /> ERRORS</Badge>}
                                        </div>
                                    </div>
                                    <FontAwesomeIcon icon={faChevronRight} className="text-muted opacity-50" />
                                </ListGroup.Item>
                            ))}
                        </ListGroup>
                    )}
                </Card.Body>
            </Card>
        </Container>
    );
};

export default TimecardView;