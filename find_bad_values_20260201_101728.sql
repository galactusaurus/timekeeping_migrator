-- SQL Queries to Find Bad Values in SQLite Database
-- Generated: 2026-02-01T10:17:28.016513
-- Use these queries to locate problematic data in your database


================================================================================
-- Query: Find Project with bad values
-- Column: Project
================================================================================
-- Project contains pipe delimiter
-- Expected pattern: \|
-- Found 49 unique bad value(s)
SELECT * FROM "TimeEntries"
WHERE "Project" = '150-39 CRAB ORCHARD' OR "Project" = '2025 OHIO FALLS MONITORING SURVEY' OR "Project" = '5252 BARDSTOWN ROAD - GAS EASEMENT Survey and Exhi' OR "Project" = 'BLOSS ROAD WATERLINE EXTENSION' OR "Project" = 'BOURBON 7 PIPELINE PROJ- PHASE 1' OR "Project" = 'BURGIN TAP REMOVAL-KYTC PERMITS' OR "Project" = 'COFFEY FARM' OR "Project" = 'COLEMAN- STONE 69KV LIDAR CORRIDOR' OR "Project" = 'CREATING NEW TRACT' OR "Project" = 'CUT OUT 4 ACRES-BUTCHERSTOWN RD' OR "Project" = 'CUT OUT 80 ACRES ON US 150' OR "Project" = 'DIVIDE FARM IN CASEY CO' OR "Project" = 'DIVIDE PROPERTY @ 775 EPHESUS SCHOOL RD' OR "Project" = 'ENGINEERING SERVICES' OR "Project" = 'EW BROWN BESS MONUMENTS' OR "Project" = 'EW BROWN LANDFILL-AERIAL LIDAR SURVEYS' OR "Project" = 'Emergency Operation Center Lot - RCIDA Park #2' OR "Project" = 'FALCON TO SALYERSVILLE FIBER-KYTC PERMITS' OR "Project" = 'FALL ROCK-MANCHESTER STRUCTURE STAKING' OR "Project" = 'FLOOD STUDY' OR "Project" = 'GENERAL ENGINEERING' OR "Project" = 'HAZARD-JACKSON STRUCTURE 60A' OR "Project" = 'HWY 150 PROPERTY IN MT VERNON' OR "Project" = 'HWY 461 MIDLAND FARMS-CONSTRUCTION STAKING' OR "Project" = 'KU PARK- ROCKY BRANCH- MICHELE LORAN' OR "Project" = 'KY 635' OR "Project" = 'KY HWY 39' OR "Project" = 'LI 16385 PR LONDON-MANCHESTER EASEMENT RESEARCH' OR "Project" = 'LI-162892 WEST CLIFF - DANVILLE' OR "Project" = 'LI-167952 -- Mercer Co Solar Line' OR "Project" = 'LI-171598 LTG LONDON-MANCHESTER 534' OR "Project" = 'MARION CO TAP PERMIT' OR "Project" = 'MARK PROPERTY LINE' OR "Project" = 'MINEOLA PIKE LINE CONSTRUCTION' OR "Project" = 'MULTI FAMILY DEVELOPMENT ON TRACT 4B' OR "Project" = 'NEW CAMP-ORINOCO-GROUND SURVEY' OR "Project" = 'NORRIS ROAD' OR "Project" = 'PR BIMBLE-LONDON EASEMENT RESEARCH' OR "Project" = 'PROJECT LI-162246 DSP-PAVILLION DRIVE' OR "Project" = 'PROPERTY LINE/HWY RIGHT OF WAY SURVEY' OR "Project" = 'RETRACE & DIVIDE 30 ACRES' OR "Project" = 'RETRACE 1036 STONEHILL CT' OR "Project" = 'RETRACE LOT IN BRODHEAD' OR "Project" = 'RETRACE LOT ON KY HWY 300-MAIN ST' OR "Project" = 'RETRACEMENT-KY HWY 1247' OR "Project" = 'SHARPS BUILDING- KY 914' OR "Project" = 'STAKE PROPERTY LINES' OR "Project" = 'TYNER - SOUTH FORK' OR "Project" = 'WOODBINE-RR & KYTC PERMITS';


================================================================================
-- Query: Review all Project values
-- Column: Project
================================================================================
-- Find all rows where Project does NOT match expected pattern
-- Pattern: \|
-- Note: This column has regex pattern: \|
-- Manual pattern matching or application-level filtering may be needed
-- Review the following for invalid values:
SELECT DISTINCT "Project", COUNT(*) as count
FROM "TimeEntries"
GROUP BY "Project"
ORDER BY count DESC;

