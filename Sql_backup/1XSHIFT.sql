-- Show me Everything
--SELECT jo.JobNum, jo.OprSeq, ldt.SetupPctComplete , ldt.LaborQty , ldt.ScrapQty, ldt.DiscrepQty, emp.Name, ldt.PayrollDate, ldt.Shift, ldt.LaborNote, ldt.LaborType, ldt.ReWork 
--FROM EpicorLive11.dbo.JobOper jo 
--INNER JOIN EpicorLive11.dbo.LaborDtl ldt ON jo.OprSeq = ldt.OprSeq AND ldt.JobNum = jo.JobNum 
--LEFT OUTER JOIN EpicorLive11.dbo.EmpBasic emp ON ldt.EmployeeNum = emp.EmpID 
--WHERE jo.JobNum = 'SD1419' AND jo.OpCode IN ('SWISS', 'CNC')
--ORDER BY ldt.ChangeDate, ldt.DspClockInTime ASC
--
--
-- To Hand Verify
-- If the total count doesn't match, we can send this list to the Cell lead for potential canditates of people who forgot to do their 1xSHIFT
-- TODO: test out later if can take the Entry day and Shift and do a difference with what MeasurLink has for the observations, and hopefully narrow
-- down our list of potential candidates to one.
--SELECT jo.JobNum, jo.OprSeq, ldt.SetupPctComplete , ldt.LaborQty , ldt.ScrapQty, ldt.DiscrepQty, emp.Name, ldt.PayrollDate , ldt.Shift, ldt.LaborNote, ldt.LaborType 
--FROM EpicorLive11.dbo.JobOper jo 
--INNER JOIN EpicorLive11.dbo.LaborDtl ldt ON jo.OprSeq = ldt.OprSeq AND ldt.JobNum = jo.JobNum
--LEFT OUTER JOIN EpicorLive11.dbo.EmpBasic emp ON ldt.EmployeeNum = emp.EmpID 
--WHERE jo.JobNum = '003697-2-5' AND jo.OpCode IN ('SWISS', 'CNC') AND (ldt.LaborQty + ldt.ScrapQty + ldt.DiscrepQty) > 0 AND ldt.LaborNote <> 'Adjustment' AND ldt.LaborNote NOT LIKE ('%DMR%')
--AND ldt.Shift IN (1,2) AND ldt.ReWork = 0
--ORDER BY ldt.PayrollDate, ldt.DspClockInTime ASC

-- DatabaseModule.Get1XSHIFTInsps()
-- Working Query to find the total count and nothing else
SELECT src.JobNum, COUNT(*) [Count]
FROM (SELECT DISTINCT jo.JobNum, ldt.PayrollDate, ldt.Shift
FROM EpicorLive11.dbo.JobOper jo 
INNER JOIN EpicorLive11.dbo.LaborDtl ldt ON jo.OprSeq = ldt.OprSeq AND ldt.JobNum = jo.JobNum 
LEFT OUTER JOIN EpicorLive11.dbo.EmpBasic emp ON ldt.EmployeeNum = emp.EmpID 
WHERE jo.JobNum = ? AND jo.OprSeq = ? AND ldt.LaborNote <> 'Adjustment' AND ldt.LaborNote NOT LIKE ('%DMR%') AND ldt.LaborQty > 0 
AND ldt.Shift IN (1,2) AND ldt.ReWork = 0) src
GROUP BY src.JobNum ;

-- DatabaseModule.Get1XSHIFTDetails()
-- Detailed list of Shifts worked from Query above. Email this information to the cell leads
SELECT jo.JobNum, emp.Name, emp.EmpID, ldt.PayrollDate, ldt.Shift, ldt.LaborQty
FROM EpicorLive11.dbo.JobOper jo 
INNER JOIN EpicorLive11.dbo.LaborDtl ldt ON jo.OprSeq = ldt.OprSeq AND ldt.JobNum = jo.JobNum 
LEFT OUTER JOIN EpicorLive11.dbo.EmpBasic emp ON ldt.EmployeeNum = emp.EmpID 
WHERE jo.JobNum = ? AND jo.OprSeq = ? AND ldt.LaborNote <> 'Adjustment' AND ldt.LaborNote NOT LIKE ('%DMR%') AND ldt.LaborQty > 0 
AND ldt.Shift IN (1,2) AND ldt.ReWork = 0;

--DatabaseModule.Get1XSHIFTTimeFrames()
-- Get the shifts with clock in and clock out time in minute from the Unix Epoch
SELECT jo.JobNum, emp.Name, emp.EmpID, ldt.ClockInDate,
	DATEDIFF(MINUTE, '1970-01-01 00:00:00', CAST(ldt.ClockInDate AS NVARCHAR) + ' 00:' + CAST(ldt.DspClockInTime AS NVARCHAR))[MinClockIn],
	ldt.ClockOutMinute - ldt.ClockInMInute[MinutesDelta]
FROM EpicorLive11.dbo.JobOper jo 
INNER JOIN EpicorLive11.dbo.LaborDtl ldt ON jo.OprSeq = ldt.OprSeq AND ldt.JobNum = jo.JobNum 
LEFT OUTER JOIN EpicorLive11.dbo.EmpBasic emp ON ldt.EmployeeNum = emp.EmpID 
WHERE jo.JobNum = ? AND jo.OprSeq = ? AND ldt.LaborNote <> 'Adjustment' AND ldt.LaborNote NOT LIKE ('%DMR%') AND ldt.LaborQty > 0 
AND ldt.Shift IN (1,2) AND ldt.ReWork = 0

