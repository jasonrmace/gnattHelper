/* global React, ReactBootstrap, Excel */
const { useState, useEffect } = React;
const { Button, Card, Badge, Spinner } = window.ReactBootstrap || {};

const ProjectList = ({ refreshTrigger }) => {
  // --- STATE ---
  const [projects, setProjects] = useState([]);
  const [isFetching, setIsFetching] = useState(false);

  // --- DATA FETCHING ---
  const fetchProjects = async () => {
    setIsFetching(true);
    try {
      await Excel.run(async (context) => {
        // 1. Fetch Team Mapping Table First
        const teamSheet = context.workbook.worksheets.getItem("Team");
        // Get the used range of the Team sheet (adjust range if you have a specific table layout)
        const teamRange = teamSheet.getUsedRange();
        teamRange.load("text");

        // 2. Find Gantt Chart Footer
        const sheet = context.workbook.worksheets.getItem("GanttChart");
        const footerRange = sheet.getRange("A:A").find("DO NOT DELETE", { completeMatch: false, matchCase: false });
        footerRange.load("rowIndex");
        
        await context.sync();

        // 3. Build Team Lookup Map (Key: Current Lead Value -> Value: "First Last")
        const teamMap = {};
        const teamRows = teamRange.text;
        
        // Skip header row (index 0) if your Team sheet has headers
        for (let i = 1; i < teamRows.length; i++) {
          const lookupKey = teamRows[i][0]?.trim(); // e.g., ShortName, ID, or First Name
          const firstName = teamRows[i][1]?.trim() || ""; // Col B
          const lastName = teamRows[i][2]?.trim() || "";  // Col C
          
          if (lookupKey) {
            teamMap[lookupKey] = `${firstName} ${lastName}`.trim();
          }
        }

        // 4. Calculate Gantt Range (Row 8 start)
        const dataStartIndex = 7;
        const footerIndex = footerRange.rowIndex;
        const rowCount = footerIndex - dataStartIndex;

        if (rowCount <= 0) {
          setProjects([]);
          return;
        }

        // 5. Get Project Data
        const dataRange = sheet.getRangeByIndexes(dataStartIndex, 0, rowCount, 8);
        dataRange.load("text");
        await context.sync();

        // 6. Filter, Map, and Cross-Reference
        const formattedData = dataRange.text
          // Filter A: Must have a Name in Col B
          .filter(row => row[1] !== "")
          // Filter B: Must be a Whole Number in Col A (Projects Only)
          .filter(row => {
            const id = parseFloat(row[0]);
            return !isNaN(id) && Number.isInteger(id);
          })
          .map((row) => {
            const rawLead = row[2]?.trim();
            // Cross-reference with the team map, fallback to raw string if not found
            const fullLeadName = teamMap[rawLead] || rawLead;

            return {
              id: row[0],       // Col A
              name: row[1],     // Col B
              lead: fullLeadName, // Cross-referenced full name
              start: row[4],    // Col E
              end: row[5],      // Col F
              percent: row[7]   // Col H
            };
          });

        setProjects(formattedData);
      });
    } catch (error) {
      console.error("Fetch error:", error);
    } finally {
      setIsFetching(false);
    }
  };

  // Auto-refresh when trigger changes
  useEffect(() => {
    fetchProjects();
  }, [refreshTrigger]);

  // --- UI RENDER ---
  return (
    <div className="mt-4">
      <div className="d-flex justify-content-between align-items-center mb-2">
        <h6 className="m-0 fw-bold text-primary">Active Projects ({projects.length})</h6>
        <Button variant="link" size="sm" className="text-decoration-none p-0" onClick={fetchProjects}>
          {isFetching ? <Spinner animation="border" size="sm" /> : "Refresh"}
        </Button>
      </div>

      {projects.length === 0 && !isFetching && (
        <div className="text-center text-muted small mt-2">No projects found.</div>
      )}

      <div style={{ maxHeight: "calc(100vh - 180px)", overflowY: "auto" }}>
        {projects.map((p, index) => (
          <Card key={index} className="mb-2 shadow-sm border-0">
            <Card.Body className="p-2">
              <div className="d-flex justify-content-between align-items-start">
                <div>
                  <Badge bg="primary" className="me-2">#{p.id}</Badge>
                  <span className="fw-bold text-dark">{p.name}</span>
                </div>
                <Badge bg={p.percent === "100%" ? "success" : p.percent === "0%" ? "danger" : "warning"} pill>
                  {p.percent || "0%"}
                </Badge>
              </div>

              <div className="mt-2 small text-muted">
                <div className="d-flex justify-content-between">
                  <span>👤 {p.lead || "Unassigned"}</span>
                </div>

                <div className="d-flex justify-content-between border-top pt-1 mt-1">
                  {p.start === "TBD" && (
                    <>
                      <span>📅 {p.start}</span>
                    </>
                  )}
                  {p.start === "" && <span>📅 TBD</span>}
                  {p.start !== "TBD" && p.start !== "" && (
                    <>
                      <span>📅 {p.start}</span>
                      <span> ➔ </span>
                      <span>{p.end}</span>
                    </>
                  )}
                </div>
              </div>
            </Card.Body>
          </Card>
        ))}
      </div>
    </div>
  );
};

window.ProjectList = ProjectList;
