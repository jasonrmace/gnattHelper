/* global React, ReactBootstrap, Excel */
import React, { useState, useEffect } from 'react';
import { Button, Card, Spinner, Table, Modal, Form, Row, Col, Badge } from 'react-bootstrap';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faUserTie, faPlus, faPencil, faTrash, faSort, faSortUp, faSortDown } from '@fortawesome/free-solid-svg-icons';
import { IdentityLogic } from '../utils/identityLogic';
import { ChangelogLogic } from '../utils/changelogLogic';

const AUTHORIZED_ADMINS = ["Rob", "Kevin", "Rob Kreps", "Kevin Rittner", "Jason", "Jason Mace"];

const SubContractorManager = ({ refreshTrigger }) => {
    const [members, setMembers] = useState([]);
    const [isLoading, setIsLoading] = useState(false);
    const [showModal, setShowModal] = useState(false);
    const [showDeleteModal, setShowDeleteModal] = useState(false);
    const [activeMember, setActiveMember] = useState(null);
    const [formData, setFormData] = useState({ 
        first: "", title: "Sub Contractor", color: "#FFFFFF"
    });
    const [isSaving, setIsSaving] = useState(false);
    const [sortConfig, setSortConfig] = useState({ key: 'first', direction: 'asc' });

    const isAdmin = AUTHORIZED_ADMINS.includes(IdentityLogic.getIdentity());

    const requestSort = (key) => {
        let direction = (sortConfig.key === key && sortConfig.direction === 'asc') ? 'desc' : 'asc';
        setSortConfig({ key, direction });
    };

    useEffect(() => {
        fetchTeam();
    }, [refreshTrigger]);

    const fetchTeam = async () => {
        setIsLoading(true);
        try {
            await Excel.run(async (context) => {
                const table = context.workbook.worksheets.getItem("Team").tables.getItem("Team");
                const range = table.getDataBodyRange();
                const columns = table.columns;
                
                columns.load("items/name");
                range.load("values");
                await context.sync();

                const colMap = {};
                columns.items.forEach((col, idx) => colMap[col.name] = idx);

                const data = range.values.map((row, idx) => ({
                    index: idx,
                    first: row[colMap["First Name"]],
                    last: row[colMap["Last Name"]],
                    title: row[colMap["Title"]],
                    office: row[colMap["Office"]],
                    manager: row[colMap["Manager"]]
                }));

                setMembers(data);
            });
        } catch (error) {
            console.error("Fetch SubContractor Error:", error);
        } finally {
            setIsLoading(false);
        }
    };

    const handleOpenModal = (member = null) => {
        if (member) {
            setActiveMember(member);
            setFormData({ 
                first: member.first, title: "Sub Contractor", 
                color: "#FFFFFF"
            });
        } else {
            setActiveMember(null);
            setFormData({ first: "", title: "Sub Contractor", color: "#FFFFFF" });
        }
        setShowModal(true);
    };

    const handleSave = async () => {
        setIsSaving(true);
        try {
            await Excel.run(async (context) => {
                const table = context.workbook.worksheets.getItem("Team").tables.getItem("Team");
                const columns = table.columns;
                columns.load("items/name");
                await context.sync();

                const colMap = {};
                columns.items.forEach((col, idx) => colMap[col.name] = idx);

                let row;
                if (activeMember) {
                    row = table.rows.getItemAt(activeMember.index);
                } else {
                    row = table.rows.add(null, [new Array(columns.items.length).fill("")]);
                }

                const rowValues = new Array(columns.items.length).fill("");
                rowValues[colMap["First Name"]] = formData.first;
                rowValues[colMap["Last Name"]] = ""; // Clear last name for subcontractors
                rowValues[colMap["Title"]] = "Sub Contractor";
                rowValues[colMap["Office"]] = ""; // Subcontractors do not get assigned an office
                rowValues[colMap["Manager"]] = ""; // Subcontractors do not get assigned a manager

                row.values = [rowValues];

                // Set the background color for the Color cell to Yellow (#FFFF00)
                const colorCell = row.getRange().getCell(0, colMap["Color"]);
                colorCell.format.fill.color = "#FFFF00";

                await ChangelogLogic.logChange(context, `${activeMember ? "Edited" : "Added"} Sub Contractor: ${formData.first}`);
                await context.sync();
            });
            setShowModal(false);
            fetchTeam();
            if (window.GlobalToast) window.GlobalToast.success("Sub Contractor saved.");
        } catch (error) {
            console.error(error);
        } finally {
            setIsSaving(false);
        }
    };

    const handleDelete = async () => {
        setIsSaving(true);
        try {
            await Excel.run(async (context) => {
                const table = context.workbook.worksheets.getItem("Team").tables.getItem("Team");
                table.rows.getItemAt(activeMember.index).delete();
                await context.sync();
            });
            setShowDeleteModal(false);
            fetchTeam();
        } catch (error) { console.error(error); }
        finally { setIsSaving(false); }
    };

    const sortItems = (items) => {
        const sortableItems = [...items];
        if (sortConfig.key !== null) {
            sortableItems.sort((a, b) => {
                let aValue = a[sortConfig.key] || "";
                let bValue = b[sortConfig.key] || "";
                if (typeof aValue === 'string') aValue = aValue.toLowerCase();
                if (typeof bValue === 'string') bValue = bValue.toLowerCase();
                if (aValue < bValue) return sortConfig.direction === 'asc' ? -1 : 1;
                if (aValue > bValue) return sortConfig.direction === 'asc' ? 1 : -1;
                return 0;
            });
        }
        return sortableItems;
    };

    const getSortIcon = (key) => {
        if (sortConfig.key !== key) return <FontAwesomeIcon icon={faSort} className="ms-1 opacity-25" />;
        return sortConfig.direction === 'asc' 
            ? <FontAwesomeIcon icon={faSortUp} className="ms-1 text-primary" />
            : <FontAwesomeIcon icon={faSortDown} className="ms-1 text-primary" />;
    };

    const sortedMembers = sortItems(members.filter(m => m.title === "Sub Contractor"));

    return (
        <div className="mt-2">
            <div className="d-flex justify-content-between align-items-center mb-4">
                <h4 className="fw-bold text-primary m-0">
                    <FontAwesomeIcon icon={faUserTie} className="me-2" />
                    Sub Contractor Management
                </h4>
                <Button variant="primary" size="sm" onClick={() => handleOpenModal()}>
                    <FontAwesomeIcon icon={faPlus} className="me-1" /> Add Contractor
                </Button>
            </div>

            <Card className="shadow-sm border-0">
                <Card.Body className="p-0">
                    {isLoading ? (
                        <div className="text-center py-5"><Spinner animation="border" variant="primary" /></div>
                    ) : (
                        <Table responsive hover className="mb-0 align-middle" style={{ fontSize: '0.85rem' }}>
                            <thead className="bg-light">
                                <tr>
                                    <th className="border-0 ps-3 cursor-pointer" onClick={() => requestSort('first')} style={{ cursor: 'pointer', userSelect: 'none' }}>
                                        Name {getSortIcon('first')}
                                    </th>
                                    <th className="border-0 text-end pe-3">Actions</th>
                                </tr>
                            </thead>
                            <tbody>
                                {sortedMembers.map((m, i) => (
                                    <tr key={i}>
                                        <td className="ps-3 fw-bold">{m.first} {m.last}</td>
                                        <td className="text-end pe-3">
                                            <Button variant="link" size="sm" className="p-0 me-2" onClick={() => handleOpenModal(m)}>
                                                <FontAwesomeIcon icon={faPencil} />
                                            </Button>
                                            {isAdmin && (
                                                <Button variant="link" size="sm" className="p-0 text-danger" onClick={() => { setActiveMember(m); setShowDeleteModal(true); }}>
                                                    <FontAwesomeIcon icon={faTrash} />
                                                </Button>
                                            )}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </Table>
                    )}
                </Card.Body>
            </Card>

            <Modal show={showModal} onHide={() => setShowModal(false)} centered>
                <Modal.Header closeButton>
                    <Modal.Title className="small fw-bold">Manage Sub Contractor</Modal.Title>
                </Modal.Header>
                <Modal.Body>
                    <Form.Group className="mb-2">
                        <Form.Label className="small fw-bold text-muted">NAME</Form.Label>
                        <Form.Control size="sm" type="text" value={formData.first} onChange={e => setFormData({...formData, first: e.target.value})} />
                    </Form.Group>
                    <Form.Group className="mb-2">
                        <Form.Label className="small fw-bold text-muted">TITLE</Form.Label>
                        <Form.Control size="sm" type="text" value="Sub Contractor" disabled />
                    </Form.Group>
                </Modal.Body>
                <Modal.Footer>
                    <Button variant="secondary" size="sm" onClick={() => setShowModal(false)}>Cancel</Button>
                    <Button variant="primary" size="sm" onClick={handleSave} disabled={isSaving || !formData.first}>
                        {isSaving ? <Spinner animation="border" size="sm" /> : "Save Contractor"}
                    </Button>
                </Modal.Footer>
            </Modal>

            <Modal show={showDeleteModal} onHide={() => setShowDeleteModal(false)} centered size="sm">
                <Modal.Header closeButton className="py-2 bg-light">
                    <Modal.Title style={{fontSize: "1rem"}} className="fw-bold text-danger">Remove Contractor?</Modal.Title>
                </Modal.Header>
                <Modal.Body className="small text-center py-4">
                    Are you sure you want to remove <br/>
                    <strong className="text-dark">{activeMember?.first} {activeMember?.last}</strong>?
                </Modal.Body>
                <Modal.Footer className="py-1 bg-light">
                    <Button variant="secondary" size="sm" onClick={() => setShowDeleteModal(false)}>Cancel</Button>
                    <Button variant="danger" size="sm" onClick={handleDelete} disabled={isSaving}>
                        {isSaving ? <Spinner animation="border" size="sm" /> : "Remove"}
                    </Button>
                </Modal.Footer>
            </Modal>
        </div>
    );
};

export default SubContractorManager;