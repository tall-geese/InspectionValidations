SELECT jh.JobNum, jh.Company, jh.PartNum, jh.RevisionNum, pr.Character01, jh.PartDescription, jh.DrawNum, 
	(SELECT MIN(jo.QtyCompleted)
	FROM EpicorLive10.dbo.JobOper jo
	WHERE jo.JobNum = ? AND jo.OpComplete = 1
	GROUP BY jo.JobNum)[Qty Complete]
FROM EpicorLive10.dbo.JobHead jh 
LEFT OUTER JOIN EpicorLive10.dbo.Project pr ON jh.ProjectID = pr.ProjectID 
WHERE jh.JobNum = ? AND jh.Company = 'JPMC'