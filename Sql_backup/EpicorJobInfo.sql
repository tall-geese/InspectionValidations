
--The Qty Completed is the smallest total qty coming from a completed operation. This should be the one before DHR

SELECT jh.JobNum, jh.Company, jh.PartNum, jh.RevisionNum, pr.Character01, jh.PartDescription, jh.DrawNum, 
    (SELECT TOP 1 SUM(ld.LaborQty)
	FROM EpicorLive10.dbo.LaborDtl ld
	INNER JOIN EpicorLive10.dbo.JobOper jo ON ld.JobNum = jo.JobNum AND ld.OprSeq = jo.OprSeq
	WHERE ld.JobNum = ? AND jo.OpComplete = 1
	GROUP BY ld.JobNum, ld.OprSeq
	ORDER BY SUM(ld.LaborQty) ASC)[Qty Complete], jh.TravelerLastPrinted
FROM EpicorLive10.dbo.JobHead jh 
LEFT OUTER JOIN EpicorLive10.dbo.Project pr ON jh.ProjectID = pr.ProjectID 
WHERE jh.JobNum = ? AND jh.Company = 'JPMC'


/*	Old Version, couldnt find qty in an operation that was filed under "rework"

SELECT jh.JobNum, jh.Company, jh.PartNum, jh.RevisionNum, pr.Character01, jh.PartDescription, jh.DrawNum, 
	(SELECT MIN(jo.QtyCompleted)
	FROM EpicorLive10.dbo.JobOper jo
	WHERE jo.JobNum = ? AND jo.OpComplete = 1
	GROUP BY jo.JobNum)[Qty Complete]
FROM EpicorLive10.dbo.JobHead jh 
LEFT OUTER JOIN EpicorLive10.dbo.Project pr ON jh.ProjectID = pr.ProjectID 
WHERE jh.JobNum = ? AND jh.Company = 'JPMC'

*/