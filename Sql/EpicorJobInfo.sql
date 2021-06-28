SELECT jh.JobNum, jh.Company, jh.PartNum, jh.RevisionNum, jo.Character01, pr.Character01, LEFT(rs.Description,LEN(rs.Description)-11)[Machine], rs.cell_c, jh.PartDescription 
FROM EpicorLive10.dbo.JobHead jh 
LEFT OUTER JOIN EpicorLive10.dbo.Project pr ON jh.ProjectID = pr.ProjectID 
LEFT OUTER JOIN EpicorLive10.dbo.JobOper jo ON jo.JobNum = jh.JobNum 
LEFT OUTER JOIN EpicorLive10.dbo.JobOpDtl jdt ON jo.JobNum = jdt.JobNum AND jo.OprSeq = jdt.OprSeq 
INNER JOIN EpicorLive10.dbo.Resource rs ON jdt.ResourceID = rs.ResourceID 
WHERE jh.JobNum = ? AND jh.Company = 'JPMC' AND jo.OpCode IN ('SWISS','CNC')