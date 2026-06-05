/* global React, ReactBootstrap, Excel */
import React, { useState, useEffect } from 'react';
import { Button, Card, Spinner, Table, Modal, Form, Row, Col, Badge } from 'react-bootstrap';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faUserGear, faPlus, faPencil, faTrash, faSyncAlt, faSort, faSortUp, faSortDown } from '@fortawesome/free-solid-svg-icons';
import { IdentityLogic } from '../utils/identityLogic';
import { ChangelogLogic } from '../utils/changelogLogic';

const TeamManager = ({ refreshTrigger }) => {
    const [members, setMembers] = useState([]);
    const [isLoading, setIsLoading] = useState(false);
    const [showModal, setShowModal] = useState(false);
    const [showDeleteModal, setShowDeleteModal] = useState(false);
    const [activeMember, setActiveMember] = useState(null);
    const [formData, setFormData] = useState({ 
        first: "", last: "", title: "", color: "#FFFFFF", office: "Houston", manager: "" 
    });
    const [isSaving, setIsSaving] = useState(false);
    const [sortConfig, setSortConfig] = useState({ key: 'first', direction: 'asc' });

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

                // We need the background color from the "Color" column
                const colorCol = columns.getItem("Color").getDataBodyRange();
                const colorProps = colorCol.getCellProperties({ format: { fill: { color: true } } });
                await context.sync();

                const data = range.values.map((row, idx) => ({
                    index: idx,
                    first: row[colMap["First Name"]],
                    last: row[colMap["Last Name"]],
                    title: row[colMap["Title"]],
                    office: row[colMap["Office"]],
                    manager: row[colMap["Manager"]],
                    color: colorProps.value[idx][0].format.fill.color
                }));

                setMembers(data);
            });
        } catch (error) {
            console.error("Fetch Team Error:", error);
        } finally {
            setIsLoading(false);
        }
    };

    const handleOpenModal = (member = null) => {
        if (member) {
            setActiveMember(member);
            setFormData({ 
                first: member.first, last: member.last, title: member.title, 
                color: member.color || "#FFFFFF", office: member.office, manager: member.manager 
            });
        } else {
            setActiveMember(null);
            setFormData({ first: "", last: "", title: "", color: "#FFFFFF", office: "Houston", manager: "" });
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
                    row = table.rows.add(null, [new Array(columns.items.length).fill("")]).getItemAt(0);
                }

                // Set values
                const rowValues = new Array(columns.items.length).fill("");
                rowValues[colMap["First Name"]] = formData.first;
                rowValues[colMap["Last Name"]] = formData.last;
                rowValues[colMap["Title"]] = formData.title;
                rowValues[colMap["Office"]] = formData.office;
                rowValues[colMap["Manager"]] = formData.manager;
                rowValues[colMap["Color"]] = ""; // Keep text empty

                row.values = [rowValues];

                // Set the background color for the Color cell
                const colorCell = row.getRange().getCell(0, colMap["Color"]);
                colorCell.format.fill.color = formData.color;

                await ChangelogLogic.logChange(context, `${activeMember ? "Edited" : "Added"} Team Member: ${formData.first} ${formData.last}`);
                await context.sync();
            });
            setShowModal(false);
            fetchTeam();
            if (window.GlobalToast) window.GlobalToast.success("Team member saved.");
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
                await ChangelogLogic.logChange(context, `Deleted Team Member: ${activeMember.first}`);
                await context.sync();
            });
            setShowDeleteModal(false);
            fetchTeam();
            if (window.GlobalToast) window.GlobalToast.error("Team member removed.");
        } catch (error) {
            console.error(error);
        } finally {
            setIsSaving(false);
        }
    };

    const sortItems = (items) => {
        const sortableItems = [...items];
        if (sortConfig.key !== null) {
            sortableItems.sort((a, b) => {
                let aValue = a[sortConfig.key] || "";
                let bValue = b[sortConfig.key] || "";

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

    const getSortIcon = (key) => {
        if (sortConfig.key !== key) return <FontAwesomeIcon icon={faSort} className="ms-1 opacity-25" />;
        return sortConfig.direction === 'asc' 
            ? <FontAwesomeIcon icon={faSortUp} className="ms-1 text-primary" />
            : <FontAwesomeIcon icon={faSortDown} className="ms-1 text-primary" />;
    };

    const filteredMembers = members.filter(m => m.first !== "General" && m.title !== "Sub Contractor");
    const sortedMembers = sortItems(filteredMembers);

    return (
        <div className="mt-2">
            <div className="d-flex justify-content-between align-items-center mb-4">
                <h4 className="fw-bold text-primary m-0">
                    <FontAwesomeIcon icon={faUserGear} className="me-2" />
                    Team Management
                </h4>
                <Button variant="primary" size="sm" onClick={() => handleOpenModal()}>
                    <FontAwesomeIcon icon={faPlus} className="me-1" /> Add Member
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
                                    <th className="border-0 ps-3 cursor-pointer" style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => requestSort('first')}>
                                        Name {getSortIcon('first')}
                                    </th>
                                    <th className="border-0 cursor-pointer" style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => requestSort('title')}>
                                        Title {getSortIcon('title')}
                                    </th>
                                    <th className="border-0 cursor-pointer" style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => requestSort('office')}>
                                        Office {getSortIcon('office')}
                                    </th>
                                    <th className="border-0 cursor-pointer" style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => requestSort('manager')}>
                                        Manager {getSortIcon('manager')}
                                    </th>
                                    <th className="border-0 text-center">Color</th>
                                    <th className="border-0 text-end pe-3">Actions</th>
                                </tr>
                            </thead>
                            <tbody>
                                {sortedMembers.map((m, i) => (
                                    <tr key={i}>
                                        <td className="ps-3 fw-bold">{m.first} {m.last}</td>
                                        <td>{m.title}</td>
                                        <td>
                                            <Badge 
                                                bg={m.office === "Houston" ? "danger" : "primary"}
                                                style={m.office === "Houston" ? { backgroundColor: "#ff4500" } : {}}
                                            >
                                                {m.office}
                                            </Badge>
                                        </td>
                                        <td>{m.manager}</td>
                                        <td className="text-center">
                                            <div style={{ width: '20px', height: '20px', backgroundColor: m.color, border: '1px solid #ddd', margin: '0 auto', borderRadius: '4px' }}></div>
                                        </td>
                                        <td className="text-end pe-3">
                                            <Button variant="link" size="sm" className="p-0 me-2" onClick={() => handleOpenModal(m)}>
                                                <FontAwesomeIcon icon={faPencil} />
                                            </Button>
                                            <Button variant="link" size="sm" className="p-0 text-danger" onClick={() => { setActiveMember(m); setShowDeleteModal(true); }}>
                                                <FontAwesomeIcon icon={faTrash} />
                                            </Button>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </Table>
                    )}
                </Card.Body>
            </Card>

            {/* ADD/EDIT MODAL */}
            <Modal show={showModal} onHide={() => setShowModal(false)} centered>
                <Modal.Header closeButton>
                    <Modal.Title className="small fw-bold">{activeMember ? "Edit Team Member" : "Add Team Member"}</Modal.Title>
                </Modal.Header>
                <Modal.Body>
                    <Row className="g-2 mb-2">
                        <Col xs={6}>
                            <Form.Group>
                                <Form.Label className="small fw-bold text-muted">FIRST NAME</Form.Label>
                                <Form.Control size="sm" type="text" value={formData.first} onChange={e => setFormData({...formData, first: e.target.value})} />
                            </Form.Group>
                        </Col>
                        <Col xs={6}>
                            <Form.Group>
                                <Form.Label className="small fw-bold text-muted">LAST NAME</Form.Label>
                                <Form.Control size="sm" type="text" value={formData.last} onChange={e => setFormData({...formData, last: e.target.value})} />
                            </Form.Group>
                        </Col>
                    </Row>
                    <Form.Group className="mb-2">
                        <Form.Label className="small fw-bold text-muted">TITLE</Form.Label>
                        <Form.Control size="sm" type="text" value={formData.title} onChange={e => setFormData({...formData, title: e.target.value})} />
                    </Form.Group>
                    <Row className="g-2 mb-2">
                        <Col xs={6}>
                            <Form.Group>
                                <Form.Label className="small fw-bold text-muted">OFFICE</Form.Label>
                                <Form.Select size="sm" value={formData.office} onChange={e => setFormData({...formData, office: e.target.value})}>
                                    <option value="Houston">Houston</option>
                                    <option value="Dallas">Dallas</option>
                                </Form.Select>
                            </Form.Group>
                        </Col>
                        <Col xs={6}>
                            <Form.Group>
                                <Form.Label className="small fw-bold text-muted">MANAGER</Form.Label>
                                <Form.Select size="sm" value={formData.manager} onChange={e => setFormData({...formData, manager: e.target.value})}>
                                    <option value="">None</option>
                                    <option value="Rob">Rob</option>
                                    <option value="Kevin">Kevin</option>
                                </Form.Select>
                            </Form.Group>
                        </Col>
                    </Row>
                    <Form.Group className="mb-2">
                        <Form.Label className="small fw-bold text-muted">CALENDAR COLOR</Form.Label>
                        <div className="d-flex align-items-center">
                            <Form.Control 
                                type="color" 
                                size="sm" 
                                value={formData.color} 
                                onChange={e => setFormData({...formData, color: e.target.value})} 
                                className="me-2"
                                style={{ width: '50px' }}
                            />
                            <Form.Text className="text-muted small">This color is used for Gantt chart task bars.</Form.Text>
                        </div>
                    </Form.Group>
                </Modal.Body>
                <Modal.Footer>
                    <Button variant="secondary" size="sm" onClick={() => setShowModal(false)}>Cancel</Button>
                    <Button variant="primary" size="sm" onClick={handleSave} disabled={isSaving || !formData.first}>
                        {isSaving ? <Spinner animation="border" size="sm" /> : "Save Member"}
                    </Button>
                </Modal.Footer>
            </Modal>

            {/* DELETE MODAL */}
            <Modal show={showDeleteModal} onHide={() => setShowDeleteModal(false)} centered size="sm">
                <Modal.Header closeButton className="py-2 bg-light">
                    <Modal.Title style={{fontSize: "1rem"}} className="fw-bold text-danger">Remove Member?</Modal.Title>
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

export default TeamManager;