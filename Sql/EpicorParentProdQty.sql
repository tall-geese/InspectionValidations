

--What is the sum of all LaborQty parts for the given operation and job?
	--Excluding when the job has been adjusted (Split off into a smaller job)
--SELECT COALESCE(SUM(ld.LaborQty), 0)
--FROM EpicorLive11.dbo.LaborDtl ld
--WHERE ld.JobNum = ? AND ld.LaborNote <> 'Adjustment' AND ld.OprSeq = (SELECT TOP 1 jo.PrimaryProdOpDtl
--																			FROM EpicorLive11.dbo.JobOper jo
--																			WHERE jo.JobNum= ?)
																			

--What is the sum of all LaborQty parts for the given operation and job?
	--Excluding when the job has been adjusted (Split off into a smaller job)
SELECT COALESCE(SUM(ld.LaborQty), 0)
FROM EpicorLive11.dbo.LaborDtl ld
WHERE ld.JobNum = ? AND ld.EmployeeNum <> 'JOBSPLIT' AND ld.OprSeq = (SELECT TOP 1 jo.PrimaryProdOpDtl
																			FROM EpicorLive11.dbo.JobOper jo
																			WHERE jo.JobNum= ?)