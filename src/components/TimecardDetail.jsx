/* global Excel */
import React, { useState, useEffect } from 'react';
import { Container, Card, Button, Row, Col, Form, Spinner, Table, Badge, Alert, Modal } from 'react-bootstrap';
import DatePicker from "react-datepicker";
import "react-datepicker/dist/react-datepicker.css";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faArrowLeft, faSave, faClock, faBriefcase, faPlane, faUmbrellaBeach, faCalculator, faExclamationTriangle } from '@fortawesome/free-solid-svg-icons';

const TimecardDetail = ({ period, onBack }) => {
    const [loading, setLoading] = useState(true);
    const [saving, setSaving] = useState(false);
    const [showSubmitModal, setShowSubmitModal] = useState(false);
    const [week1, setWeek1] = useState([]);
    const [week2, setWeek2] = useState([]);
    const [summary, setSummary] = useState({
        w1Total: 0, w1Office: 0, w1Travel: 0, w1Pto: 0, w1Check: 0,
        w2Total: 0, w2Office: 0, w2Travel: 0, w2Pto: 0, w2Check: 0,
        grandTotal: 0, grandOfficePto: 0, grandTravel: 0, grandCheck: 0,
        regOffPto: 0, otOffice: 0, regTravel: 0, otTravel: 0,
        totalDetail: 0, finalCheck: 0
    });

    const PTO_TYPES = ["Vacation", "Sick", "Personal", "Holiday", "Jury Duty", "Bereavement"];

    const loadSheetData = async () => {
        setLoading(true);
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(period.name);
                
                // Week 1: Rows 7-17, Columns C-I (Index 2-8)
                const week1Range = sheet.getRangeByIndexes(6, 2, 11, 7);
                // Week 2: Rows 19-29, Columns C-I (Index 2-8)
                const week2Range = sheet.getRangeByIndexes(18, 2, 11, 7);

                // Summary Totals from Column K (Index 10)
                // Week 1: K13-K17 (Idx 12-16)
                const w1Summary = sheet.getRangeByIndexes(12, 10, 5, 1);
                // Week 2: K25-K29 (Idx 24-28)
                const w2Summary = sheet.getRangeByIndexes(24, 10, 5, 1);
                // Combined: K31-K34 (Idx 30-33)
                const combinedSummary = sheet.getRangeByIndexes(30, 10, 4, 1);
                // Column H Summary (Idx 7): H32-H38 (Idx 31-37)
                const hSummary = sheet.getRangeByIndexes(31, 7, 7, 1);

                week1Range.load("text");
                week2Range.load("text");
                w1Summary.load("values");
                w2Summary.load("values");
                combinedSummary.load("values");
                hSummary.load("values");
                await context.sync();

                const mapData = (data) => {
                    const days = [];
                    for (let col = 0; col < 7; col++) {
                        days.push({
                            date: data[0][col],
                            timeIn: data[1][col],
                            lunchOut: data[2][col],
                            lunchIn: data[3][col],
                            timeOut: data[4][col],
                            total: data[6][col], // Row 13 / 25
                            office: data[7][col], // Row 14 / 26
                            travel: data[8][col], // Row 15 / 27
                            pto: data[9][col],    // Row 16 / 28
                            ptoType: data[10][col] // Row 17 / 29
                        });
                    }
                    return days;
                };

                setWeek1(mapData(week1Range.text));
                setWeek2(mapData(week2Range.text));
                
                setSummary({
                    w1Total: w1Summary.values[0][0],
                    w1Office: w1Summary.values[1][0],
                    w1Travel: w1Summary.values[2][0],
                    w1Pto: w1Summary.values[3][0],
                    w1Check: w1Summary.values[4][0],
                    w2Total: w2Summary.values[0][0],
                    w2Office: w2Summary.values[1][0],
                    w2Travel: w2Summary.values[2][0],
                    w2Pto: w2Summary.values[3][0],
                    w2Check: w2Summary.values[4][0],
                    grandTotal: combinedSummary.values[0][0],
                    grandOfficePto: combinedSummary.values[1][0],
                    grandTravel: combinedSummary.values[2][0],
                    grandCheck: combinedSummary.values[3][0],
                    regOffPto: hSummary.values[0][0],
                    otOffice: hSummary.values[1][0],
                    regTravel: hSummary.values[2][0],
                    otTravel: hSummary.values[3][0],
                    totalDetail: hSummary.values[5][0],
                    finalCheck: hSummary.values[6][0]
                });
            });
        } catch (error) {
            console.error("Error loading sheet details:", error);
        } finally {
            setLoading(false);
        }
    };

    /**
     * Converts Excel time values/strings to JS Date objects for the picker.
     */
    const parseToDate = (val) => {
        if (!val || val === "TBD") return null;
        try {
            const date = new Date(`1970-01-01 ${val}`);
            return isNaN(date.getTime()) ? null : date;
        } catch (e) { return null; }
    };

    const updateExcelCell = async (rowIndex, colIndex, value, key, isWeek2 = false) => {
        if (period.isSubmitted) return;
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(period.name);
                const cell = sheet.getCell(rowIndex, colIndex + 2); // +2 for Column C offset
                
                // Format time objects back to strings for Excel if needed, or clear if null
                if (value instanceof Date) {
                    value = value.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                } else if (value === null) {
                    value = "";
                }

                cell.values = [[value]];
                
                // If updating a time, also fetch the "Total" cell (calculated by Excel formula)
                const totalRowOffset = isWeek2 ? 24 : 12; // Row 13 (idx 12) or Row 25 (idx 24)
                const totalCell = sheet.getCell(totalRowOffset, colIndex + 2);
                totalCell.load("text");

                // Load full summaries to update footer
                const w1Sum = sheet.getRangeByIndexes(12, 10, 5, 1);
                const w2Sum = sheet.getRangeByIndexes(24, 10, 5, 1);
                const grandSum = sheet.getRangeByIndexes(30, 10, 4, 1);
                const hSum = sheet.getRangeByIndexes(31, 7, 7, 1);

                w1Sum.load("values");
                w2Sum.load("values");
                grandSum.load("values");
                hSum.load("values");

                await context.sync();

                // Update local state so UI reflects the change and the new total immediately
                const updateState = (prev) => {
                    const newWeek = [...prev];
                    newWeek[colIndex] = { 
                        ...newWeek[colIndex], 
                        [key]: value,
                        total: totalCell.text[0][0]
                    };
                    return newWeek;
                };

                if (isWeek2) setWeek2(updateState);
                else setWeek1(updateState);

                setSummary({
                    w1Total: w1Sum.values[0][0],
                    w1Office: w1Sum.values[1][0],
                    w1Travel: w1Sum.values[2][0],
                    w1Pto: w1Sum.values[3][0],
                    w1Check: w1Sum.values[4][0],
                    w2Total: w2Sum.values[0][0],
                    w2Office: w2Sum.values[1][0],
                    w2Travel: w2Sum.values[2][0],
                    w2Pto: w2Sum.values[3][0],
                    w2Check: w2Sum.values[4][0],
                    grandTotal: grandSum.values[0][0],
                    grandOfficePto: grandSum.values[1][0],
                    grandTravel: grandSum.values[2][0],
                    grandCheck: grandSum.values[3][0],
                    regOffPto: hSum.values[0][0],
                    otOffice: hSum.values[1][0],
                    regTravel: hSum.values[2][0],
                    otTravel: hSum.values[3][0],
                    totalDetail: hSum.values[5][0],
                    finalCheck: hSum.values[6][0]
                });
            });
        } catch (error) {
            console.error("Write error:", error);
        }
    };

    useEffect(() => {
        loadSheetData();
    }, [period.name]);

    const getDayInfo = (dateStr) => {
        if (!dateStr || dateStr === "TBD") return { label: dateStr, isOffDay: false };
        // Parse Excel date string (e.g., "6/1/2026")
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return { label: dateStr, isOffDay: false };
        const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        return {
            label: `${dayNames[d.getDay()]} ${d.getMonth() + 1}/${d.getDate()}`,
            isOffDay: d.getDay() === 0 || d.getDay() === 6 // Sunday=0, Saturday=6
        };
    };

    // --- VALIDATION HELPER ---
    const checkWeekErrors = (days) => {
        return days.some(d => {
            const total = parseFloat(d.total) || 0;
            const allocSum = (parseFloat(d.office) || 0) + (parseFloat(d.travel) || 0) + (parseFloat(d.pto) || 0);
            const hasMismatch = total > 0 && Math.abs(allocSum - total) > 0.01;
            const ptoHours = parseFloat(d.pto) || 0;
            const ptoError = ptoHours > 0 && !d.ptoType;
            return total < 0 || hasMismatch || ptoError;
        });
    };

    const errorsExist = summary.w1Check !== 0 || summary.w2Check !== 0 || summary.grandCheck !== 0 || summary.finalCheck !== 0 || checkWeekErrors(week1) || checkWeekErrors(week2);

    const renderWeekTable = (days, startRowIndex) => (
        <div className="table-responsive">
            <style>{`
                .react-datepicker-wrapper { width: 100%; }
                
                /* Ensure column background shows through inputs and selects */
                .off-day-cell {
                    background-color: #eeeeee !important; /* Visible light grey */
                }
                .off-day-cell input, 
                .off-day-cell select,
                .off-day-cell .react-datepicker-wrapper,
                .off-day-cell .react-datepicker__input-container input {
                    background: transparent !important;
                }

                /* Hide dropdown arrow for locked/disabled PTO Type selects */
                select.form-select:disabled {
                    background-image: none !important;
                }

                .time-picker-custom {
                    width: 100%;
                    border: none;
                    text-align: center;
                    font-size: 0.75rem;
                    padding: 4px;
                }
                .time-picker-custom:focus {
                    outline: 1px solid #0d6efd;
                    background-color: #fff;
                }
                .react-datepicker__time-container { width: 100px !important; }
                .react-datepicker__time-box { width: 100px !important; }
                
                /* Style for the clear icon */
                .react-datepicker__close-icon {
                    background-color: transparent !important;
                    padding: 0 !important;
                    right: 4px !important;
                    display: flex !important;
                    align-items: center !important;
                }
                .react-datepicker__close-icon::after {
                    content: "" !important;
                    display: block !important;
                    width: 14px !important;
                    height: 14px !important;
                    background-image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" fill="%236c757d"><path d="M256 512A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM175 175c9.4-9.4 24.6-9.4 33.9 0l47 47 47-47c9.4-9.4 24.6-9.4 33.9 0s9.4 24.6 0 33.9l-47 47 47 47c9.4 9.4 9.4 24.6 0 33.9s-24.6 9.4-33.9 0l-47-47-47 47c-9.4 9.4-24.6 9.4-33.9 0s-9.4-24.6 0-33.9l47-47-47-47c-9.4-9.4-9.4-24.6 0-33.9z"/></svg>') !important;
                    background-repeat: no-repeat !important;
                    background-position: center !important;
                    background-color: transparent !important;
                    border-radius: 0 !important;
                }
                .react-datepicker__close-icon:hover::after {
                    background-image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" fill="%23dc3545"><path d="M256 512A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM175 175c9.4-9.4 24.6-9.4 33.9 0l47 47 47-47c9.4-9.4 24.6-9.4 33.9 0s9.4 24.6 0 33.9l-47 47 47 47c9.4 9.4 9.4 24.6 0 33.9s-24.6 9.4-33.9 0l-47-47-47 47c-9.4 9.4-24.6 9.4-33.9 0s-9.4-24.6 0-33.9l47-47-47-47c-9.4-9.4-9.4-24.6 0-33.9z"/></svg>') !important;
                }
            `}</style>
            <Table bordered size="sm" className="text-center align-middle bg-white shadow-sm" style={{ fontSize: '0.75rem', minWidth: '600px' }}>
                <thead className="bg-light">
                    <tr>
                        <th style={{ width: '100px' }}>Type</th>
                        {days.map((d, i) => {
                            const info = getDayInfo(d.date);
                            return (
                                <th key={i} className={info.isOffDay ? "off-day-cell" : ""}>
                                    {info.label || `Day ${i+1}`}
                                </th>
                            );
                        })}
                    </tr>
                </thead>
                <tbody>
                    {/* TIME ENTRIES */}
                    <tr className="table-secondary-subtle fw-bold text-uppercase" style={{ fontSize: '0.65rem' }}>
                        <td className="text-start ps-2 py-1" colSpan={8}><FontAwesomeIcon icon={faClock} className="me-2"/>Time Tracking</td>
                    </tr>
                    {['timeIn', 'lunchOut', 'lunchIn', 'timeOut'].map((key, rowOffset) => (
                        <tr key={key}>
                            <td className="text-start ps-2 fw-bold text-muted text-capitalize">{key.replace(/([A-Z])/g, ' $1')}</td>
                            {days.map((d, colIdx) => {
                                const info = getDayInfo(d.date);
                                return (
                                <td key={colIdx} className={`p-0 ${info.isOffDay ? "off-day-cell" : ""}`}>
                                    <DatePicker
                                        selected={parseToDate(d[key])}
                                        onChange={(date) => updateExcelCell(startRowIndex + 1 + rowOffset, colIdx, date, key, startRowIndex > 10)}
                                        showTimeSelect
                                        showTimeSelectOnly
                                        timeIntervals={15}
                                        timeCaption="Time"
                                        dateFormat="h:mm aa"
                                        disabled={period.isSubmitted}
                                        className="time-picker-custom"
                                        isClearable={!period.isSubmitted}
                                        clearButtonTitle="Clear This Time"
                                    />
                                </td>
                                );
                            })}
                        </tr>
                    ))}
                    <tr className="table-light fw-bold">
                        <td className="text-start ps-2">Daily Total</td>
                        {days.map((d, i) => {
                            const isNegative = parseFloat(d.total) < 0;
                            const info = getDayInfo(d.date);
                            return (
                                <td key={i} className={`${isNegative ? "bg-danger-subtle text-danger" : ""} ${info.isOffDay && !isNegative ? "off-day-cell" : ""}`}>
                                    {d.total}
                                </td>
                            );
                        })}
                    </tr>

                    {/* ALLOCATIONS */}
                    <tr className="table-primary-subtle fw-bold text-uppercase" style={{ fontSize: '0.65rem' }}>
                        <td className="text-start ps-2 py-1" colSpan={8}><FontAwesomeIcon icon={faBriefcase} className="me-2"/>Allocations</td>
                    </tr>
                    {[
                        { label: 'Office/Site', key: 'office', icon: faBriefcase, row: 7 },
                        { label: 'Travel', key: 'travel', icon: faPlane, row: 8 },
                        { label: 'PTO Hours', key: 'pto', icon: faUmbrellaBeach, row: 9 }
                    ].map((rowDef) => (
                        <tr key={rowDef.key}>
                            <td className="text-start ps-2 fw-bold text-muted">{rowDef.label}</td>
                            {days.map((d, colIdx) => {
                                const info = getDayInfo(d.date);
                                const totalVal = parseFloat(d.total) || 0;
                                const allocSum = (parseFloat(d.office) || 0) + (parseFloat(d.travel) || 0) + (parseFloat(d.pto) || 0);
                                const hasMismatch = totalVal > 0 && Math.abs(allocSum - totalVal) > 0.01;
                                
                                return (
                                    <td key={colIdx} className={`p-0 ${info.isOffDay ? "off-day-cell" : ""}`}>
                                        <input 
                                            type="number" 
                                            className={`form-control form-control-sm border-0 text-center rounded-0 p-1 ${hasMismatch ? "bg-warning-subtle" : ""}`}
                                            defaultValue={d[rowDef.key]}
                                            disabled={period.isSubmitted}
                                            onBlur={(e) => updateExcelCell(startRowIndex + rowDef.row, colIdx, e.target.value, rowDef.key, startRowIndex > 10)}
                                        />
                                    </td>
                                );
                            })}
                        </tr>
                    ))}
                    <tr>
                        <td className="text-start ps-2 fw-bold text-muted">PTO Type</td>
                        {days.map((d, colIdx) => {
                            const info = getDayInfo(d.date);
                            const ptoHours = parseFloat(d.pto) || 0;
                            const ptoError = ptoHours > 0 && !d.ptoType;
                            
                            return (
                                <td key={colIdx} className={`p-0 ${info.isOffDay ? "off-day-cell" : ""}`}>
                                    <select 
                                        className={`form-select form-select-sm border-0 text-center rounded-0 p-1 ${ptoError ? "bg-danger-subtle" : ""}`}
                                        defaultValue={d.ptoType}
                                        disabled={period.isSubmitted}
                                        onChange={(e) => updateExcelCell(startRowIndex + 10, colIdx, e.target.value, 'ptoType', startRowIndex > 10)}
                                        style={{ fontSize: '0.65rem' }}
                                    >
                                        <option value=""></option>
                                        {PTO_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                                    </select>
                                </td>
                            );
                        })}
                    </tr>
                </tbody>
            </Table>
        </div>
    );

    return (
        <Container className="mt-3 px-2">
            {/* HEADER */}
            <Card className="mb-3 border-0 shadow-sm">
                <Card.Body className="p-2 d-flex justify-content-between align-items-center bg-primary text-white rounded">
                    <div className="d-flex align-items-center">
                        <Button variant="link" className="text-white p-0 me-3" onClick={onBack}>
                            <FontAwesomeIcon icon={faArrowLeft} />
                        </Button>
                        <div>
                            <h6 className="mb-0 fw-bold">{period.start} — {period.end}</h6>
                            {/* <small className="opacity-75">{period.name}</small> */}
                        </div>
                    </div>
                    {period.isSubmitted ? (
                        <Badge bg="success" pill>SUBMITTED (READ ONLY)</Badge>
                    ) : (
                        <div className="d-flex align-items-center gap-2">
                            <Badge bg="warning" text="dark" pill>EDITING MODE</Badge>
                            <Button variant="light" size="sm" className="fw-bold py-0" style={{ fontSize: '0.7rem' }} onClick={() => setShowSubmitModal(true)}>
                                Submit Timecard
                            </Button>
                        </div>
                    )}
                </Card.Body>
            </Card>

            {loading ? (
                <div className="text-center py-5">
                    <Spinner animation="grow" variant="primary" />
                    <p className="text-muted mt-2">Opening timesheet...</p>
                </div>
            ) : (
                <>
                    {/* Checksum Errors */}
                    {(summary.w1Check !== 0 || summary.w2Check !== 0 || summary.grandCheck !== 0 || summary.finalCheck !== 0) && (
                        <div className="p-3 border-top">
                            <Alert variant="danger" className="mb-0 py-2 small d-flex align-items-center">
                                <FontAwesomeIcon icon={faExclamationTriangle} className="me-2" />
                                <div>
                                    <div className="fw-bold">Pay Period Errors Detected</div>
                                    {summary.w1Check !== 0 && <div>Week 1 Allocation Mismatch: {summary.w1Check}</div>}
                                    {summary.w2Check !== 0 && <div>Week 2 Allocation Mismatch: {summary.w2Check}</div>}
                                    {summary.finalCheck !== 0 && <div>Grand Total Detail Mismatch: {summary.finalCheck}</div>}
                                </div>
                            </Alert>
                        </div>
                    )}
                    
                    {/* WEEK 1 */}
                    <h6 className="fw-bold text-muted small mb-2 ps-1">WEEK ONE</h6>
                    {renderWeekTable(week1, 6)}

                    {/* WEEK 2 */}
                    <h6 className="fw-bold text-muted small mt-4 mb-2 ps-1">WEEK TWO</h6>
                    {renderWeekTable(week2, 18)}

                    {/* PAY PERIOD SUMMARY */}
                    <Card className="mt-4 shadow-sm border-0">
                        <Card.Header className="bg-light py-2">
                            <h6 className="mb-0 fw-bold text-primary small">
                                <FontAwesomeIcon icon={faCalculator} className="me-2" />
                                Pay Period Summary
                            </h6>
                        </Card.Header>
                        <Card.Body className="p-0">
                            <Table responsive borderless size="sm" className="mb-0 text-center small">
                                <thead className="border-bottom bg-light-subtle">
                                    <tr>
                                        <th className="text-start ps-3">Category</th>
                                        <th>Week 1</th>
                                        <th>Week 2</th>
                                        <th className="bg-primary-subtle">Total</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr className="border-bottom">
                                        <td className="text-start ps-3 fw-bold">Total Hours</td>
                                        <td>{summary.w1Total}</td>
                                        <td>{summary.w2Total}</td>
                                        <td className="fw-bold">{summary.grandTotal}</td>
                                    </tr>
                                    <tr>
                                        <td className="text-start ps-3">Office / PTO</td>
                                        <td>{Number(summary.w1Office) + Number(summary.w1Pto)}</td>
                                        <td>{Number(summary.w2Office) + Number(summary.w2Pto)}</td>
                                        <td className="fw-bold">{summary.grandOfficePto}</td>
                                    </tr>
                                    <tr>
                                        <td className="text-start ps-3">Travel</td>
                                        <td>{summary.w1Travel}</td>
                                        <td>{summary.w2Travel}</td>
                                        <td className="fw-bold">{summary.grandTravel}</td>
                                    </tr>
                                    <tr className="border-top">
                                        <td className="text-start ps-3">Regular Office + PTO</td>
                                        <td colSpan={2}></td>
                                        <td className="fw-bold">{summary.regOffPto}</td>
                                    </tr>
                                    <tr>
                                        <td className="text-start ps-3">Overtime Office</td>
                                        <td colSpan={2}></td>
                                        <td className="fw-bold">{summary.otOffice}</td>
                                    </tr>
                                    <tr>
                                        <td className="text-start ps-3">Regular Travel</td>
                                        <td colSpan={2}></td>
                                        <td className="fw-bold">{summary.regTravel}</td>
                                    </tr>
                                    <tr>
                                        <td className="text-start ps-3">Overtime Travel</td>
                                        <td colSpan={2}></td>
                                        <td className="fw-bold">{summary.otTravel}</td>
                                    </tr>
                                    <tr className="border-top">
                                        <td className="text-start ps-3 fw-bold">Total Detail</td>
                                        <td colSpan={2}></td>
                                        <td className="fw-bold">{summary.totalDetail}</td>
                                    </tr>
                                </tbody>
                            </Table>
                        </Card.Body>
                    </Card>

                    <div className="mt-4 pb-4">
                        <Card className="bg-light border-0">
                            <Card.Body className="p-3 text-center small text-muted italic">
                                {period.isSubmitted 
                                    ? "This timesheet is submitted and locked. Contact HR or your Manager to un-submit."
                                    : "Changes are automatically saved to Excel as you type (or when you leave the input field)."
                                }
                            </Card.Body>
                        </Card>
                    </div>
                </>
            )}

            {/* SUBMIT MODAL */}
            <Modal show={showSubmitModal} onHide={() => setShowSubmitModal(false)} centered>
                <Modal.Header closeButton>
                    <Modal.Title className="text-primary fw-bold" style={{ fontSize: '1.1rem' }}>Submit Timesheet</Modal.Title>
                </Modal.Header>
                <Modal.Body>
                    {errorsExist ? (
                        <Alert variant="danger">
                            <FontAwesomeIcon icon={faExclamationTriangle} className="me-2" />
                            <strong>Cannot Submit:</strong> Your timesheet contains errors or unallocated hours. Please resolve all warnings and checksum errors before submitting.
                        </Alert>
                    ) : (
                        <>
                            <p>Are you sure you want to submit this timesheet? This will lock the sheet for editing and finalize your hours for payroll processing.</p>
                            <p className="mb-0"><small><i>By clicking submit, you agree that all the times worked are accurate and you have submitted all PTO (and it has been approved).</i></small></p>
                        </>
                    )}
                </Modal.Body>
                <Modal.Footer className="bg-light py-2">
                    <Button variant="secondary" size="sm" onClick={() => setShowSubmitModal(false)}>Cancel</Button>
                    <Button variant="primary" size="sm" disabled={errorsExist} onClick={() => setShowSubmitModal(false)}>
                        Confirm Submission
                    </Button>
                </Modal.Footer>
            </Modal>
        </Container>
    );
};

export default TimecardDetail;