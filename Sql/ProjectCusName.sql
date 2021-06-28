
SELECT pr.Character01
FROM EpicorLive10.dbo.JobHead jh 
LEFT OUTER JOIN EpicorLive10.dbo.Project pr ON pr.ProjectID = jh.ProjectID 
WHERE jh.JobNum = ?