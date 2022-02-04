
SELECT jh.JobNum, jo.Character01, LEFT(rs.Description,
								CASE 
									WHEN (LEN(rs.Description) - 11) < 3 THEN LEN(rs.Description)
									ELSE (LEN(rs.Description) - 11)
								END)[Machine], rs.cell_c, jo.OprSeq, jo.OpCode 
FROM EpicorLive10.dbo.JobHead jh 
LEFT OUTER JOIN EpicorLive10.dbo.JobOper jo ON jo.JobNum = jh.JobNum 
LEFT OUTER JOIN EpicorLive10.dbo.JobOpDtl jdt ON jo.JobNum = jdt.JobNum AND jo.OprSeq = jdt.OprSeq 
INNER JOIN EpicorLive10.dbo.Resource rs ON jdt.ResourceID = rs.ResourceID 
WHERE jh.JobNum = ? AND jo.OpCode IN ('SWISS','CNC') AND rs.cell_c  <> ''