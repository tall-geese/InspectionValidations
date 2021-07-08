-- Show me Everything
--SELECT jo.JobNum, jo.OprSeq, ldt.SetupPctComplete , ldt.LaborQty , ldt.ScrapQty, ldt.DiscrepQty, emp.Name, ldt.PayrollDate, ldt.Shift, ldt.LaborNote, ldt.LaborType, ldt.ReWork 
--FROM EpicorLive10.dbo.JobOper jo 
--INNER JOIN EpicorLive10.dbo.LaborDtl ldt ON jo.OprSeq = ldt.OprSeq AND ldt.JobNum = jo.JobNum 
--LEFT OUTER JOIN EpicorLive10.dbo.EmpBasic emp ON ldt.EmployeeNum = emp.EmpID 
--WHERE jo.JobNum = 'SD1419' AND jo.OpCode IN ('SWISS', 'CNC')
--ORDER BY ldt.ChangeDate, ldt.DspClockInTime ASC
--
--
-- To Hand Verify
-- If the total count doesn't match, we can send this list to the Cell lead for potential canditates of people who forgot to do their 1xSHIFT
-- TODO: test out later if can take the Entry day and Shift and do a difference with what MeasurLink has for the observations, and hopefully narrow
-- down our list of potential candidates to one.
--SELECT jo.JobNum, jo.OprSeq, ldt.SetupPctComplete , ldt.LaborQty , ldt.ScrapQty, ldt.DiscrepQty, emp.Name, ldt.PayrollDate , ldt.Shift, ldt.LaborNote, ldt.LaborType 
--FROM EpicorLive10.dbo.JobOper jo 
--INNER JOIN EpicorLive10.dbo.LaborDtl ldt ON jo.OprSeq = ldt.OprSeq AND ldt.JobNum = jo.JobNum
--LEFT OUTER JOIN EpicorLive10.dbo.EmpBasic emp ON ldt.EmployeeNum = emp.EmpID 
--WHERE jo.JobNum = '003697-2-5' AND jo.OpCode IN ('SWISS', 'CNC') AND (ldt.LaborQty + ldt.ScrapQty + ldt.DiscrepQty) > 0 AND ldt.LaborNote <> 'Adjustment' AND ldt.LaborNote NOT LIKE ('%DMR%')
--AND ldt.Shift IN (1,2) AND ldt.ReWork = 0
--ORDER BY ldt.PayrollDate, ldt.DspClockInTime ASC

-- Working Query to find the total count and nothing else
SELECT src.JobNum, COUNT(*) [Count]
FROM (SELECT DISTINCT jo.JobNum, ldt.PayrollDate, ldt.Shift
FROM EpicorLive10.dbo.JobOper jo 
INNER JOIN EpicorLive10.dbo.LaborDtl ldt ON jo.OprSeq = ldt.OprSeq AND ldt.JobNum = jo.JobNum 
LEFT OUTER JOIN EpicorLive10.dbo.EmpBasic emp ON ldt.EmployeeNum = emp.EmpID 
WHERE jo.JobNum = ? AND jo.OpCode IN ('SWISS', 'CNC') AND ldt.LaborNote <> 'Adjustment' AND ldt.LaborNote NOT LIKE ('%DMR%') AND ldt.LaborQty > 0 
AND ldt.Shift IN (1,2) AND ldt.ReWork = 0) src
GROUP BY src.JobNum