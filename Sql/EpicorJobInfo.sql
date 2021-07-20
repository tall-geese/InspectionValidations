SELECT jh.JobNum, jh.Company, jh.PartNum, jh.RevisionNum, pr.Character01, jh.PartDescription, jh.DrawNum 
FROM EpicorLive10.dbo.JobHead jh 
LEFT OUTER JOIN EpicorLive10.dbo.Project pr ON jh.ProjectID = pr.ProjectID 
LEFT OUTER JOIN EpicorLive10.dbo.JobOper jo ON jo.JobNum = jh.JobNum 
WHERE jh.JobNum = ? AND jh.Company = 'JPMC'