/* global React, ReactDOM, Office, ReactBootstrap, Excel */
import React, { useState, useEffect, useRef } from 'react';
import ReactDOM from 'react-dom/client';
import { Container, Alert, Spinner } from 'react-bootstrap';
import toast, { Toaster } from 'react-hot-toast';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faBell, faTimes, faCrosshairs, faUserCircle, faCalendarDays } from '@fortawesome/free-solid-svg-icons';

import { IdentityLogic } from './utils/identityLogic';
import { EventListeners } from './utils/eventListeners';
import { ChangelogLogic } from './utils/changelogLogic';
import { TimecardLogic } from './utils/timecardLogic.jsx';
import ProjectList from './components/ProjectList';
import PTOManager from './components/PTOManager';
import HomePage from './components/HomePage';
import CreatePTO from './components/CreatePTO';
import TeamManager from './components/TeamManager';
import SubContractorManager from './components/SubContractorManager';
import CreateProject from './components/CreateProject';
// Assuming these will also be converted to ES modules soon:
import MainNavbar from './components/MainNavbar';
import SettingsPage from './components/SettingsPage';
import IdentityModal from './components/IdentityModal';
import UpdatesPage from './components/UpdatesPage';
import TimecardView from './components/TimecardView';
import CreateTimecardModal from './components/CreateTimecardModal';
import BarbizonSpinner from './components/BarbizonSpinner';
import OverdueTasks from './components/OverdueTasks';
import TasksPage from './components/TasksPage';

// 1. CONFIGURATION
const GANTT_FILENAMES = ["Barbizon Texas Project Management.xlsx"];
const DEPRECATED_FILENAMES = ["Houston Summer 2026 [Macros].xlsm", "Houston Summer 2026.xlsx"];
const GANTT_SHEET_NAMES = ["Houston", "Dallas"];
const TIMECARD_PREFIX = "2026_Timecard_Template_";
const PM_TIMELOG_PREFIX = "PM_TIMELOG_";

// --- GLOBAL LOADING OVERLAY ---
const LoadingOverlay = ({ isVisible, message }) => {
    if (!isVisible) return null;
    
    return (
        <div style={{
            position: 'fixed', top: 0, left: 0, width: '100%', height: '100%', // Ensure full screen coverage
            backgroundColor: 'rgba(255, 255, 255, 0.9)', zIndex: 9999,
            background: '#090d16',
            // backgorundColor: '#090d16 !important', zIndex: 9999,
            display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center'
        }}>
            <div className="mb-4" aria-label="Loading" style={{
                backgroundColor: 'rgba(9, 13, 22, 0.9)',
            }}>
                <BarbizonSpinner 
                    size={100} 
                    speed="1.2s" 
                    message={
                        <span className="text-center d-block">
                            {message} 
                            <br/>
                            <span className="small" style={{color: '#FFFFFF', opacity: 0.8}}>Please wait while Excel updates...</span>
                        </span>
                    } 
                />
            </div>
        </div>
    );
};

function App() {
    // --- STATE ---
    const [version] = useState("v6.3.0"); 
    const [activeTab, setActiveTab] = useState("Home");
    const [isValidFile, setIsValidFile] = useState(false);
    const [currentName, setCurrentName] = useState("");
    const [hudText, setHudText] = useState("Ready");
    const [refreshTrigger, setRefreshTrigger] = useState(0);
    const [fileError, setFileError] = useState("");
    const [highlightId, setHighlightId] = useState(null);
    const [navResetTrigger, setNavResetTrigger] = useState(0);
    const [fileType, setFileType] = useState(null); // 'gantt' or 'timecard'
    const [unseenCount, setUnseenCount] = useState(0);
    const [overdueCount, setOverdueCount] = useState(0);
    const [userOverdueCount, setUserOverdueCount] = useState(0);
    
    // LOADER & MODAL STATE
    const [isLoading, setIsLoading] = useState(false);
    const [loadingMsg, setLoadingMsg] = useState("Processing...");
    const [showIdentityModal, setShowIdentityModal] = useState(false);
    const [showCreateTimecardModal, setShowCreateTimecardModal] = useState(false);
    
    const processingRef = useRef(false);
    const timecardCheckRef = useRef(false);

    // --- HELPER: TRIGGER TOAST ---
    const showToast = (title, msg) => {
        // react-hot-toast doesn't have a direct 'title' concept like Bootstrap.
        // We'll use the message as the primary content and the title as an ID
        // to prevent duplicate toasts with the same "title" if they appear quickly.
        toast.success(msg, { id: title, duration: 6000 });
    };

    // --- HELPER: TRIGGER LIST REFRESH ---
    const triggerRefresh = () => {
        // We add a small delay to allow Excel to finalize worksheet collection 
        // updates before we attempt to re-read them in our components.
        console.log("Triggering Refresh (1s delay)...");
        setTimeout(() => {
            setRefreshTrigger(prev => prev + 1);
                }, 1000);
    };

    const fetchOverdueData = async () => {
        try {
            return await Excel.run(async (context) => {
                const sheets = ["Houston", "Dallas"];
                const currentUser = IdentityLogic.getIdentity();
                const today = new Date();
                today.setHours(0, 0, 0, 0);
                let total = 0;
                let yours = 0;

                for (const name of sheets) {
                    const sheet = context.workbook.worksheets.getItemOrNullObject(name);
                    const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
                    footerRange.load(["rowIndex", "isNullObject"]);
                    sheet.load("isNullObject");
                    await context.sync();

                    if (sheet.isNullObject || footerRange.isNullObject) continue;

                    const rowCount = footerRange.rowIndex - 7;
                    if (rowCount <= 0) continue;

                    const range = sheet.getRangeByIndexes(7, 0, rowCount, 8);
                    range.load("text");
                    await context.sync();

                    range.text.forEach(row => {
                        const idNum = parseFloat(row[0]);
                        if (!isNaN(idNum) && !Number.isInteger(idNum)) {
                            const taskEnd = new Date(row[5]);
                            if (row[5] !== "TBD" && !isNaN(taskEnd) && taskEnd < today && row[7] !== "100%") {
                                total++;
                                if (row[2] === currentUser) yours++;
                            }
                        }
                    });
                }
                setOverdueCount(total);
                setUserOverdueCount(yours);
                return { total, yours };
            });
        } catch (e) { console.warn("Overdue fetch failed", e); return { total: 0, yours: 0 }; }
    };

    const fetchUnseenCount = async (clearFirst = false) => {
        if (clearFirst) toast.dismiss();

        try {
            await Excel.run(async (context) => {
                const user = IdentityLogic.getIdentity();
                if (!user) return;

                const adminUsers = ["Rob", "Kevin", "Rob Kreps", "Kevin Rittner"];
                const isAdmin = adminUsers.includes(user);

                // 1. Update the Badge Count
                const count = await ChangelogLogic.getUnseenCount(context);
                setUnseenCount(count);

                // 2. Fetch specific logs for toasts
                const adminLogs = await ChangelogLogic.getUnseenAdminLogs(context);
                const ptoLogs = isAdmin ? await ChangelogLogic.getUnseenPTOLogs(context) : [];

                // 3. Show individual toasts for Admin updates (Global broadcasts)
                adminLogs.forEach(log => {
                    toast.custom((t) => (
                        <div className={`${t.visible ? 'animate-enter' : 'animate-leave'} bg-white shadow-lg rounded-lg pointer-events-auto d-flex border border-info`}
                             style={{ minWidth: '350px' }}>
                            <div className="flex-grow-1 p-4">
                                <div className="d-flex align-items-start h-100">
                                    <FontAwesomeIcon icon={faBell} className="text-info mt-1" />
                                    <div className="ms-3">
                                        <p className="text-sm font-bold text-dark mb-1" style={{ fontSize: '0.9rem' }}>
                                            New Administrative Update
                                        </p>
                                        <p className="text-sm text-muted mb-0" style={{ fontSize: '0.85rem' }}>
                                            {log.change}
                                        </p>
                                    </div>
                                </div>
                            </div>
                            <div className="d-flex border-start border-light align-items-stretch">
                                <button
                                    onClick={async () => {
                                        await ChangelogLogic.markAsSeen(log.change, log.timestamp);
                                        toast.dismiss(t.id);
                                        fetchUnseenCount();
                                    }}
                                    className="btn btn-link text-decoration-none border-0 px-4 py-0 d-flex align-items-center justify-content-center text-sm font-bold text-info hover:bg-light focus:outline-none"
                                    style={{ borderRadius: '0 8px 8px 0', fontSize: '0.85rem' }}
                                >
                                    Close
                                </button>
                            </div>
                        </div>
                    ), { duration: Infinity, id: `admin-update-${log.timestamp}` });
                });

                // 4. Show summary toast for standard updates
                // We subtract ptoLogs because those are handled separately with Approve/Deny buttons
                const trackedCount = adminLogs.length + ptoLogs.length;
                const nonTrackedCount = count - trackedCount;

                if (nonTrackedCount > 0) {
                    toast.custom((t) => (
                        <div className={`${t.visible ? 'animate-enter' : 'animate-leave'} bg-white shadow-lg rounded-lg pointer-events-auto d-flex border border-light`}
                             style={{ minWidth: '350px' }}>
                            <div className="flex-grow-1 p-4">
                                <div className="d-flex align-items-start h-100">
                                    <FontAwesomeIcon icon={faBell} className="text-primary mt-1" />
                                    <div className="ms-3">
                                        <p className="text-sm font-medium text-dark mb-1" style={{ fontSize: '0.9rem' }}>
                                            You have missed {nonTrackedCount} {nonTrackedCount === 1 ? 'update' : 'updates'}.
                                        </p>
                                        <button 
                                            className="btn btn-link p-0 border-0 fw-bold text-primary"
                                            style={{ textDecoration: 'underline', fontSize: '0.85rem' }}
                                            onClick={() => {
                                                setActiveTab("Updates");
                                                toast.dismiss(t.id);
                                            }}
                                        >
                                            View history here
                                        </button>
                                    </div>
                                </div>
                            </div>
                            <div className="d-flex border-start border-light align-items-stretch">
                                <button
                                    onClick={() => toast.dismiss(t.id)}
                                    className="btn btn-link text-decoration-none border-0 px-4 py-0 d-flex align-items-center justify-content-center text-sm font-bold text-muted hover:bg-light focus:outline-none"
                                    style={{ borderRadius: '0 8px 8px 0', fontSize: '0.85rem' }}
                                >
                                    Close
                                </button>
                            </div>
                        </div>
                    ), { 
                        duration: 20000,
                        id: 'unseen-summary'
                    });
                }

                // 5. Show PTO Approve/Deny Toasts (Managers only)
                if (isAdmin && ptoLogs.length > 0) {
                    ptoLogs.forEach(log => {
                        toast.custom((t) => (
                            <div className={`${t.visible ? 'animate-enter' : 'animate-leave'} bg-white shadow-lg rounded-lg pointer-events-auto d-flex border border-warning`}
                                 style={{ minWidth: '350px' }}>
                                <div className="flex-grow-1 p-4">
                                    <div className="d-flex align-items-start">
                                        <FontAwesomeIcon icon={faCalendarDays} className="text-warning mt-1" />
                                        <div className="ms-3">
                                            <p className="text-sm font-bold text-dark mb-1" style={{ fontSize: '0.9rem' }}>
                                                PTO Request: {log.author.toLowerCase().replace(/\b\w/g, s => s.toUpperCase())}
                                            </p>
                                            <p className="text-sm text-muted mb-0" style={{ fontSize: '0.85rem' }}>
                                                {log.change.replace(/Added PTO for ([^:]+)/, (match, name) => 
                                                    `Added PTO for ${name.toLowerCase().replace(/\b\w/g, s => s.toUpperCase())}`
                                                )}
                                            </p>
                                        </div>
                                    </div>
                                </div>
                                <div className="d-flex border-start border-light align-items-stretch">
                                    <div className="d-flex flex-column border-start">
                                        <button
                                            onClick={() => {
                                                setActiveTab("PTOManager");
                                                toast.dismiss(t.id);
                                            }}
                                            className="btn btn-link text-decoration-none border-bottom px-3 py-2 d-flex align-items-center justify-content-center text-sm font-bold text-primary hover:bg-light"
                                            style={{ fontSize: '0.75rem' }}
                                        >
                                            View
                                        </button>
                                        <button
                                            onClick={async () => {
                                                await ChangelogLogic.markAsSeen(log.change, log.timestamp);
                                                toast.dismiss(t.id);
                                                fetchUnseenCount();
                                                window.GlobalToast.success("PTO Approved");
                                            }}
                                            className="btn btn-link text-decoration-none border-bottom px-3 py-2 d-flex align-items-center justify-content-center text-sm font-bold text-success hover:bg-light"
                                            style={{ fontSize: '0.75rem' }}
                                        >
                                            Approve
                                        </button>
                                        <button
                                            onClick={async () => {
                                                await handleDenyPTO(log);
                                                toast.dismiss(t.id);
                                            }}
                                            className="btn btn-link text-decoration-none px-3 py-2 d-flex align-items-center justify-content-center text-sm font-bold text-danger hover:bg-light"
                                            style={{ fontSize: '0.75rem' }}
                                        >
                                            Deny
                                        </button>
                                    </div>
                                </div>
                            </div>
                        ), { duration: Infinity, id: `pto-update-${log.timestamp}` });
                    });
                }
            });
        } catch (e) { console.error(e); }
    };

    const handleDenyPTO = async (log) => {
        window.GlobalLoader.show("Denying Request...");
        try {
            await Excel.run(async (context) => {
                const table = context.workbook.tables.getItem("Vacations");
                const range = table.getDataBodyRange();
                range.load("values");
                await context.sync();

                // Description format: "Added PTO for [User]: [Start] to [End]"
                // We extract the name and start date to find the specific row
                const parts = log.change.split(":");
                const namePart = parts[0].replace("Added PTO for ", "").trim();
                const datePart = parts[1].split(" to ")[0].trim();

                const rows = range.values;
                let rowIndexToDelete = -1;

                for (let i = 0; i < rows.length; i++) {
                    if (rows[i][0] === namePart && rows[i][1] === datePart) {
                        rowIndexToDelete = i;
                        break;
                    }
                }

                if (rowIndexToDelete !== -1) {
                    table.rows.getItemAt(rowIndexToDelete).delete();
                }

                await ChangelogLogic.markAsSeen(log.change, log.timestamp);
                await context.sync();
                triggerRefresh();
                window.GlobalToast.error(`PTO for ${namePart} denied and removed.`);
            });
        } catch (e) { console.error(e); }
        finally { window.GlobalLoader.hide(); }
    };

    // --- EXPOSE LOADER TO WINDOW ---
    useEffect(() => {
        window.GlobalLoader = {
            show: (msg = "Updating...") => {
                setLoadingMsg(msg);
                setIsLoading(true);
            },
            hide: () => {
                setIsLoading(false);
            }
        };

        // Bridge for EventListeners to update UI
        window.RefreshBadge = () => {
            fetchUnseenCount();
        };

        // --- EXPOSE TOAST TO WINDOW ---
        window.GlobalToast = {
            success: (msg, opts) => toast.success(msg, opts),
            error: (msg, opts) => toast.error(msg, opts),
            info: (msg, opts) => toast(msg, { icon: '🔄', ...opts }),
        };

        // Expose refresh globally so utility files can trigger UI updates
        window.GlobalRefresh = () => {
            triggerRefresh();
        };
    }, []);

    // --- INITIALIZATION & EVENTS ---
    useEffect(() => {
        Office.onReady(async (info) => {
            if (info.host === Office.HostType.Excel) {
                console.log("Office Ready");

                const url = Office.context.document.url;
                
                // Handle New/Unsaved File
                if (!url) {
                    setIsValidFile(false);
                    setFileError("You have loaded a new excel file and Barbizon Helpers do not work here.");
                    setHudText("Inactive");
                    return;
                }

                const decodedUrl = decodeURI(url);
                // Strips query parameters (common in Web/SharePoint URLs) then splits by forward slash / or backslash \.
                const fileName = decodedUrl.split('?')[0].split(/[\\\/]/).pop();
                setCurrentName(fileName);

                const isDeprecated = DEPRECATED_FILENAMES.includes(fileName);
                const isGantt = GANTT_FILENAMES.includes(fileName);
                const isTimecard = fileName.startsWith(TIMECARD_PREFIX);
                const isPmTimelog = fileName.startsWith(PM_TIMELOG_PREFIX);
                const isAllowed = isGantt || isTimecard || isPmTimelog;

                setIsValidFile(isAllowed);

                // A. Register HUD Selection Listener
                Office.context.document.addHandlerAsync(
                    Office.EventType.DocumentSelectionChanged,
                    handleSelectionChange
                );

                if (isDeprecated) {
                    setFileError("This file is no longer supported. Please use 'Barbizon Texas Project Management.xlsx' instead.");
                    setHudText("Deprecated");
                    return;
                }

                if (!isAllowed) {
                    setFileError(`Helpers are locked for "${fileName}". Please open an approved file.`);
                    setHudText("Locked");
                    return;
                }

                if (isGantt) {
                    setFileType('gantt');

                    // Initial Load: Activate "Splash" sheet if it exists
                    Excel.run(async (context) => {
                        const splash = context.workbook.worksheets.getItemOrNullObject("Splash");
                        splash.load("isNullObject");
                        await context.sync();
                        if (!splash.isNullObject) splash.activate();
                        await context.sync();
                    }).catch(() => {});

                    // B. Register Watchdog
                    await EventListeners.register();

                    await fetchUnseenCount();
                } else if (isTimecard) {
                    setFileType('timecard');
                    setActiveTab("TimecardDashboard");
                    
                    // Trigger the auto-generation check for next period (only once)
                    if (!timecardCheckRef.current) {
                        timecardCheckRef.current = true;
                        // The logic returns true if a new sheet was actually created
                        TimecardLogic.checkAndGenerateNextPeriod().then(created => {
                            if (created) {
                                triggerRefresh();
                            }
                        });
                    }
                } else if (isPmTimelog) {
                    setFileType('pm_timelog');
                    setActiveTab("PmTimelogDashboard");
                }

                // C. CHECK IDENTITY
                const user = IdentityLogic.getIdentity();
                if (!user) {
                    setShowIdentityModal(true);
                } else {
                    // Enhanced Welcome Toast with Overdue Logic
                    fetchOverdueData().then(counts => {
                        toast.custom((t) => (
                            <div className={`${t.visible ? 'animate-enter' : 'animate-leave'} bg-white shadow-lg rounded-lg pointer-events-auto d-flex border border-primary`} style={{ minWidth: '350px', cursor: 'pointer' }} onClick={() => { setActiveTab("OverdueTasks"); toast.dismiss(t.id); }}>
                                <div className="flex-grow-1 p-3">
                                    <div className="d-flex align-items-start">
                                        <div className="ms-3">
                                            <p className="text-sm font-bold text-dark mb-1" style={{ fontSize: '0.9rem' }}>Good to see you, {user}!</p>
                                            <p className="text-sm text-muted mb-0" style={{ fontSize: '0.85rem' }}>
                                                There are {counts.total} overdue tasks, {counts.yours} are yours. <br/>
                                                <span className="text-primary fw-bold" style={{ textDecoration: 'underline' }}>Click here to view these.</span>
                                            </p>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        ), { duration: 10000 });
                    });
                }
            }
        });

        return () => {
            Office.context.document.removeHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                handleSelectionChange
            );
        };
    }, []);

    // --- MODAL HANDLER ---
    const handleModalClose = () => {
        setShowIdentityModal(false);
        // After the modal closes, check if a user was successfully saved.
        // The modal itself will trigger a reload if successful, so this path
        // is primarily for when the modal might close without a save (e.g., dev environment).
        const user = IdentityLogic.getIdentity();
        if (user) {
            showToast("Identity Saved", `Welcome to the team, ${user}!`);
        }
    };
    // --- HUD LOGIC ---
    const handleSelectionChange = async () => {
        if (processingRef.current) return;
        processingRef.current = true;
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.load("name");
                const range = context.workbook.getSelectedRange();
                range.load(["rowIndex", "rowCount"]);
                await context.sync();

                // Only run on Gantt sheets
                if (range.rowCount > 1 || range.rowIndex < 7 || !GANTT_SHEET_NAMES.includes(sheet.name)) {
                    processingRef.current = false;
                    return;
                }
                
                const currentRowIndex = range.rowIndex;
                const idCell = sheet.getRangeByIndexes(currentRowIndex, 0, 1, 1);   
                const nameCell = sheet.getRangeByIndexes(currentRowIndex, 1, 1, 1); 
                idCell.load("text"); nameCell.load("text"); 
                await context.sync();

                const currentID = idCell.text[0][0];
                const currentName = nameCell.text[0][0];
                let finalText = `${sheet.name.toUpperCase()} | Row: ${currentRowIndex + 1}`;

                if (currentID) {
                    const idNum = parseFloat(currentID);
                    if (Number.isInteger(idNum)) {
                        finalText = `PROJECT ${currentID}: ${currentName}`;
                    } else if (!isNaN(idNum)) {
                        const parentID = Math.floor(idNum).toString();
                        const foundRange = sheet.getRange("A:A").find(parentID, { 
                            completeMatch: true, matchCase: false 
                        });
                        const parentNameRange = foundRange.getOffsetRange(0, 1);
                        parentNameRange.load("text");
                        await context.sync();
                        
                        const pName = parentNameRange.text[0][0];
                        finalText = `PROJECT ${parentID}: ${pName} - Task ${currentID}`;
                    }
                }
                setHudText(finalText);
                
                const shape = sheet.shapes.getItem("TaskHUD");
                shape.load("name"); 
                await context.sync().then(() => {
                    shape.textFrame.textRange.text = finalText;
                    return context.sync();
                }).catch(() => {});
            });
        } catch (error) {
             // console.error("HUD Error:", error); 
        } finally {
            setTimeout(() => { processingRef.current = false; }, 200);
        }
    };

    // --- UI RENDER ---
    return (
        <div className="d-flex flex-column vh-100 bg-white position-relative">
            
            <Toaster 
                position="bottom-right" 
                containerStyle={{ bottom: 50, right: 20 }} 
                toastOptions={{ duration: 6000 }}
            />

            {/* GLOBAL LOADER */}
            <LoadingOverlay isVisible={isLoading} message={loadingMsg} />

            {/* IDENTITY MODAL */}
            <IdentityModal 
                show={showIdentityModal} 
                onHide={handleModalClose} 
            />

                <CreateTimecardModal 
                    show={showCreateTimecardModal}
                    onHide={() => setShowCreateTimecardModal(false)}
                    onCreated={() => triggerRefresh()}
                />

            {/* 1. HEADER */}
            <div className="flex-shrink-0">
                <MainNavbar 
                    activeTab={activeTab} 
                    setActiveTab={setActiveTab} 
                    isFileValid={isValidFile} 
                    unseenCount={unseenCount}
                    overdueCount={overdueCount}
                    fileType={fileType}
                    onAddTimecard={() => setShowCreateTimecardModal(true)}
                    onUpdatesViewed={() => setUnseenCount(0)}
                    onClearHighlight={() => setHighlightId(null)}
                    onNavReset={() => {
                        setHighlightId(null);
                        setNavResetTrigger(prev => prev + 1);
                    }}
                />
            </div>

            {/* 2. BODY */}
            <div className={`flex-grow-1 p-3 ${activeTab === "ProjectList" ? "overflow-hidden" : "overflow-auto"}`}>
                {!isValidFile ? (
                    <div className="mt-5 px-2">
                        <Alert variant="warning" className="shadow-sm border-warning text-center py-4">
                            <FontAwesomeIcon icon={faTimes} size="2x" className="text-danger mb-3" />
                            <h5 className="fw-bold">Functionality Locked</h5>
                            <p className="mb-0 small text-muted">
                                {fileError}
                            </p>
                        </Alert>
                    </div>
                ) : (
                    <>
                            {activeTab === "Home" && <HomePage setActiveTab={setActiveTab} />}
                        {activeTab === "HoustonList" && (
                            <ProjectList 
                                key={`Houston-${navResetTrigger}`}
                                sheetName="Houston" 
                                refreshTrigger={refreshTrigger} 
                                highlightId={highlightId} 
                                onClearHighlight={() => setHighlightId(null)} 
                            />
                        )}
                        {activeTab === "DallasList" && (
                            <ProjectList 
                                key={`Dallas-${navResetTrigger}`}
                                sheetName="Dallas" 
                                refreshTrigger={refreshTrigger} 
                                highlightId={highlightId} 
                                onClearHighlight={() => setHighlightId(null)} 
                            />
                        )}
                        {activeTab === "AddProject" && <CreateProject onProjectCreated={handleProjectCreated} />}
                        {activeTab === "PTOManager" && <PTOManager refreshTrigger={refreshTrigger} onNavigateToAdd={() => setActiveTab("AddPTO")} />}
                        {activeTab === "SubContractorManager" && <SubContractorManager refreshTrigger={refreshTrigger} />}
                        {activeTab === "TeamManager" && <TeamManager refreshTrigger={refreshTrigger} />}
                        {activeTab === "AddPTO" && <CreatePTO onPTOCreated={triggerRefresh} />}
                        {activeTab === "OverdueTasks" && (
                            <OverdueTasks 
                                onNavigate={(loc, taskId) => {
                                    setHighlightId(taskId);
                                    setNavResetTrigger(prev => prev + 1);
                                    setActiveTab(loc === "Houston" ? "HoustonList" : "DallasList");
                                }} 
                            />)}
                        {activeTab === "Tasks" && (
                            <TasksPage 
                                refreshTrigger={refreshTrigger}
                            />
                        )}
                        {activeTab === "TimecardDashboard" && <TimecardView currentFileName={currentName} refreshTrigger={refreshTrigger} />}
                        {activeTab === "PmTimelogDashboard" && (
                            <div className="text-center mt-5">
                                <h4 className="fw-bold text-primary">Project Management Timelog Functions</h4>
                                <p className="text-muted">Coming soon...</p>
                            </div>
                        )}
                        {activeTab === "Settings" && <SettingsPage currentFileName={currentName} />} 
                        {activeTab === "Updates" && <UpdatesPage onMarkedSeen={() => setUnseenCount(0)} />}
                    </>
                )}
            </div>

            {/* 3. FOOTER */}
            <div className="bg-primary text-white shadow-lg px-3 py-2 d-flex justify-content-between align-items-center flex-shrink-0" 
                 style={{ fontSize: "0.8rem", borderTop: "3px solid #0d6efd", zIndex: 1030 }}>

                {fileType === 'gantt' ? (
                    <span 
                        className="fw-bold text-truncate" 
                        style={{maxWidth: "80%", cursor: "help"}} 
                        title={hudText} // Tooltip for the HUD text
                    >
                        <FontAwesomeIcon icon={faCrosshairs} className="me-2 opacity-50" /> {hudText}
                    </span>
                ) : <div />}

                <span className="opacity-50" style={{fontSize: "0.7rem"}}>{version}</span>
            </div>
        </div>
        );
    }
    
const container = document.getElementById('root');
if (!window.reactRoot) {
    window.reactRoot = ReactDOM.createRoot(container);
}
window.reactRoot.render(<App />);
