/* global React, ReactBootstrap, Excel */
import React, { useState, useEffect } from 'react';
import { Button, Card, Badge, Spinner, Table, Modal, Form, Row, Col, Alert } from 'react-bootstrap';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faCalendarAlt, faTrash, faPencil, faPlus, faUser, faCalendarDays, faSyncAlt, faSort, faSortUp, faSortDown } from '@fortawesome/free-solid-svg-icons';

import { FormattingLogic } from '../utils/formattingLogic_v2';
import { VisualLogic } from '../utils/visualLogic';
import { ChangelogLogic } from '../utils/changelogLogic';
import { IdentityLogic } from '../utils/identityLogic';

const PTOManager = ({ onNavigateToAdd, refreshTrigger }) => {
    const [ptoList, setPtoList] = useState([]);
    const [isFetching, setIsFetching] = useState(false);
    const [currentUser, setCurrentUser] = useState("");
    const [isAdmin, setIsAdmin] = useState(false);

    // Edit Modal State
    const [showEdit, setShowEdit] = useState(false);
    const [editingItem, setEditingItem] = useState(null);
    const [editData, setEditData] = useState({ startDate: "", endDate: "" });
    const [isSaving, setIsSaving] = useState(false);
    const [sortConfig, setSortConfig] = useState({ key: 'startDate', direction: 'desc' });
    
    // Delete Modal State
    const [showDelete, setShowDelete] = useState(false);
    const [itemToDelete, setItemToDelete] = useState(null);

    const adminUsers = ["Rob", "Kevin", "Rob Kreps", "Kevin Rittner"];

    useEffect(() => {
        const user = IdentityLogic.getIdentity();
        setCurrentUser(user);
        setIsAdmin(adminUsers.includes(user));
        fetchPTO();
    }, [refreshTrigger]);

    const fetchPTO = async () => {
        setIsFetching(true);
        try {
            await Excel.run(async (context) => {
                const table = context.workbook.tables.getItemOrNullObject("Vacations");
                table.load("isNullObject");
                await context.sync();

                if (table.isNullObject) {
                    setPtoList([]);
                    return;
                }

                const range = table.getDataBodyRange();
                range.load("values");
                await context.sync();

                // Map rows to objects with table index for easy reference during update/delete
                const items = range.values.map((row, index) => ({
                    index: index,
                    who: row[0],
                    startDate: row[1],
                    totalDays: row[2],
                    endDate: row[3] // Calculated by Excel formula
                }));

                setPtoList(items);
            });
        } catch (error) {
            console.error("Fetch PTO Error:", error);
        } finally {
            setIsFetching(false);
        }
    };

    const handleDeleteRequest = (item) => {
        setItemToDelete(item);
        setShowDelete(true);
    };

    const handleConfirmDelete = async () => {
        if (!itemToDelete) return;
        setIsSaving(true);

        try {
            await Excel.run(async (context) => {
                const table = context.workbook.tables.getItem("Vacations");
                table.rows.getItemAt(itemToDelete.index).delete();
                
                await ChangelogLogic.logChange(context, `Deleted PTO for ${itemToDelete.who} starting ${itemToDelete.startDate}`);
                await context.sync();

                for (const name of ["Houston", "Dallas"]) {
                    await FormattingLogic.generateSmartRules(context, name);
                    await VisualLogic.refreshGridAlerts(context, name);
                }
            });
            setShowDelete(false);
            setItemToDelete(null);
            fetchPTO();
        } catch (error) { 
            console.error("Delete Error:", error);
        } finally { setIsSaving(false); }
    };

    const formatDateForDisplay = (val) => {
        if (!val) return "-";
        let d;
        if (typeof val === 'number') {
            // Convert Excel serial date to JS Date object
            d = new Date(Math.round((val - 25569) * 86400 * 1000));
        } else {
            d = new Date(val);
        }
        if (isNaN(d.getTime())) return val;
        const month = String(d.getUTCMonth() + 1).padStart(2, '0');
        const day = String(d.getUTCDate()).padStart(2, '0');
        const year = d.getUTCFullYear();
        return `${month}/${day}/${year}`;
    };

    const formatDateForInput = (val) => {
        if (!val) return "";
        let d;
        if (typeof val === 'number') {
            d = new Date(Math.round((val - 25569) * 86400 * 1000));
        } else {
            d = new Date(val);
        }
        if (isNaN(d.getTime())) return "";
        return d.toISOString().split('T')[0];
    };

    const openEditModal = (item) => {
        setEditingItem(item);
        setEditData({ 
            startDate: formatDateForInput(item.startDate), 
            endDate: formatDateForInput(item.endDate)
        });
        setShowEdit(true);
    };

    const handleSaveEdit = async () => {
        setIsSaving(true);
        try {
            await Excel.run(async (context) => {
                const table = context.workbook.tables.getItem("Vacations");
                const row = table.rows.getItemAt(editingItem.index);
                
                const start = new Date(editData.startDate);
                const end = new Date(editData.endDate);
                const duration = Math.ceil(Math.abs(end - start) / (1000 * 60 * 60 * 24)) + 1;

                // Update: Who, Start, Duration
                row.values = [[editingItem.who, editData.startDate, duration, ""]];
                
                await ChangelogLogic.logChange(context, `Edited PTO for ${editingItem.who}: New range ${editData.startDate} to ${editData.endDate}`);
                await context.sync();

                for (const name of ["Houston", "Dallas"]) {
                    await FormattingLogic.generateSmartRules(context, name);
                    await VisualLogic.refreshGridAlerts(context, name);
                }
            });
            setShowEdit(false);
            fetchPTO();
        } catch (error) { console.error(error); }
        finally { setIsSaving(false); }
    };

    const requestSort = (key) => {
        let direction = 'asc';
        if (sortConfig.key === key && sortConfig.direction === 'asc') {
            direction = 'desc';
        }
        setSortConfig({ key, direction });
    };

    const sortItems = (items) => {
        const sortableItems = [...items];
        if (sortConfig.key !== null) {
            sortableItems.sort((a, b) => {
                let aValue = a[sortConfig.key];
                let bValue = b[sortConfig.key];

                // Handle nulls
                if (aValue === null || aValue === undefined) return 1;
                if (bValue === null || bValue === undefined) return -1;

                // String normalization for sorting
                if (typeof aValue === 'string') aValue = aValue.toLowerCase();
                if (typeof bValue === 'string') bValue = bValue.toLowerCase();

                if (aValue < bValue) return sortConfig.direction === 'asc' ? -1 : 1;
                if (aValue > bValue) return sortConfig.direction === 'asc' ? 1 : -1;
                return 0;
            });
        }
        return sortableItems;
    };

    const sortedPto = sortItems(ptoList);
    const myPto = sortedPto.filter(p => p.who === currentUser);
    const othersPto = sortedPto.filter(p => p.who !== currentUser);

    return (
        <div className="mt-2">
            <div className="d-flex justify-content-between align-items-center mb-4">
                <h4 className="fw-bold text-primary m-0">
                    <FontAwesomeIcon icon={faCalendarAlt} className="me-2" />
                    PTO / Vacation Manager
                </h4>
                <Button variant="primary" size="sm" onClick={onNavigateToAdd}>
                    <FontAwesomeIcon icon={faPlus} className="me-1" /> Add PTO
                </Button>
            </div>

            {/* SECTION: MY PTO */}
            <Card className="shadow-sm mb-4 border-0 bg-light">
                <Card.Header className="bg-white border-bottom-0 pt-3">
                    <h6 className="fw-bold text-dark m-0">Your PTO Requests</h6>
                </Card.Header>
                <Card.Body className="p-0">
                    <PTOTable 
                        items={myPto} 
                        onEdit={openEditModal} 
                        onDelete={handleDeleteRequest} 
                        formatDate={formatDateForDisplay}
                        showUser={false}
                        sortConfig={sortConfig}
                        onSort={requestSort}
                        isFetching={isFetching}
                    />
                </Card.Body>
            </Card>

            {/* SECTION: TEAM PTO (Admins Only) */}
            {isAdmin && (
                <Card className="shadow-sm border-0">
                    <Card.Header className="bg-white border-bottom-0 pt-3 d-flex justify-content-between">
                        <h6 className="fw-bold text-muted m-0">All Other Team PTO</h6>
                        <Badge bg="info" text="dark">Admin Access</Badge>
                    </Card.Header>
                    <Card.Body className="p-0">
                        <PTOTable 
                            items={othersPto} 
                            onEdit={openEditModal} 
                            onDelete={handleDeleteRequest} 
                            formatDate={formatDateForDisplay}
                            showUser={true}
                            sortConfig={sortConfig}
                            onSort={requestSort}
                            isFetching={isFetching}
                        />
                    </Card.Body>
                </Card>
            )}

            {/* EDIT MODAL */}
            <Modal show={showEdit} onHide={() => setShowEdit(false)} centered>
                <Modal.Header closeButton>
                    <Modal.Title className="small fw-bold">Edit PTO: {editingItem?.who}</Modal.Title>
                </Modal.Header>
                <Modal.Body>
                    <Row className="g-2">
                        <Col xs={6}>
                            <Form.Group>
                                <Form.Label className="small fw-bold">START DATE</Form.Label>
                                <Form.Control size="sm" type="date" value={editData.startDate} onChange={e => setEditData({...editData, startDate: e.target.value})} />
                            </Form.Group>
                        </Col>
                        <Col xs={6}>
                            <Form.Group>
                                <Form.Label className="small fw-bold">END DATE</Form.Label>
                                <Form.Control size="sm" type="date" value={editData.endDate} onChange={e => setEditData({...editData, endDate: e.target.value})} />
                            </Form.Group>
                        </Col>
                    </Row>
                    <Alert variant="warning" className="mt-3 p-2 small mb-0">
                        Updating this will immediately refresh all Gantt Chart bars.
                    </Alert>
                </Modal.Body>
                <Modal.Footer>
                    <Button variant="secondary" size="sm" onClick={() => setShowEdit(false)}>Cancel</Button>
                    <Button variant="primary" size="sm" onClick={handleSaveEdit} disabled={isSaving || !editData.endDate}>
                        {isSaving ? <Spinner animation="border" size="sm" /> : "Save Changes"}
                    </Button>
                </Modal.Footer>
            </Modal>

            {/* DELETE CONFIRMATION MODAL */}
            <Modal show={showDelete} onHide={() => setShowDelete(false)} centered size="sm">
                <Modal.Header closeButton className="py-2 bg-light">
                    <Modal.Title style={{fontSize: "1rem"}} className="fw-bold text-danger">Delete PTO?</Modal.Title>
                </Modal.Header>
                <Modal.Body className="small text-center py-4">
                    Are you sure you want to delete PTO for <br/>
                    <strong className="text-dark">{itemToDelete?.who}</strong>?
                </Modal.Body>
                <Modal.Footer className="py-1 bg-light">
                    <Button variant="secondary" size="sm" onClick={() => setShowDelete(false)}>Cancel</Button>
                    <Button variant="danger" size="sm" onClick={handleConfirmDelete} disabled={isSaving}>
                        {isSaving ? <Spinner animation="border" size="sm" /> : "Yes, Delete"}
                    </Button>
                </Modal.Footer>
            </Modal>
        </div>
    );
};

const PTOTable = ({ items, onEdit, onDelete, formatDate, isFetching, sortConfig, onSort, showUser = true }) => {
    if (isFetching && items.length === 0) return <div className="text-center py-4"><Spinner animation="border" size="sm" /></div>;
    if (items.length === 0) return <div className="text-center py-4 text-muted small">No PTO entries found.</div>;

    const getSortIcon = (key) => {
        if (sortConfig.key !== key) return <FontAwesomeIcon icon={faSort} className="ms-1 opacity-25" />;
        return sortConfig.direction === 'asc' 
            ? <FontAwesomeIcon icon={faSortUp} className="ms-1 text-primary" />
            : <FontAwesomeIcon icon={faSortDown} className="ms-1 text-primary" />;
    };

    return (
        <Table responsive hover className="mb-0 align-middle" style={{ fontSize: '0.85rem' }}>
            <thead className="bg-light">
                <tr>
                    {showUser && (
                        <th className="border-0 ps-3 cursor-pointer" style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => onSort('who')}>
                            User {getSortIcon('who')}
                        </th>
                    )}
                    <th className={`border-0 cursor-pointer ${!showUser ? 'ps-3' : ''}`} style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => onSort('startDate')}>
                        Start {getSortIcon('startDate')}
                    </th>
                    <th className="border-0 cursor-pointer" style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => onSort('endDate')}>
                        End {getSortIcon('endDate')}
                    </th>
                    <th className="border-0 text-center cursor-pointer" style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => onSort('totalDays')}>
                        Days {getSortIcon('totalDays')}
                    </th>
                    <th className="border-0 text-end pe-3">Actions</th>
                </tr>
            </thead>
            <tbody>
                {items.map((item, i) => (
                    <tr key={i}>
                        {showUser && <td className="ps-3 fw-bold">{item.who}</td>}
                        <td className={!showUser ? 'ps-3' : ''}>{formatDate(item.startDate)}</td>
                        <td>{formatDate(item.endDate)}</td>
                        <td className="text-center"><Badge bg="secondary" pill>{item.totalDays}</Badge></td>
                        <td className="text-end pe-3">
                            <Button variant="link" size="sm" className="p-0 me-2 text-primary" onClick={() => onEdit(item)}>
                                <FontAwesomeIcon icon={faPencil} />
                            </Button>
                            <Button variant="link" size="sm" className="p-0 text-danger" onClick={() => onDelete(item)}>
                                <FontAwesomeIcon icon={faTrash} />
                            </Button>
                        </td>
                    </tr>
                ))}
            </tbody>
        </Table>
    );
};

export default PTOManager;