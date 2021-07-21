
SELECT jh.JobNum, jo.Character01, LEFT(rs.Description,LEN(rs.Description)-11)[Machine], rs.cell_c, src.ProdQty, jo.OprSeq, jo.OpCode 
FROM EpicorLive10.dbo.JobHead jh 
LEFT OUTER JOIN EpicorLive10.dbo.JobOper jo ON jo.JobNum = jh.JobNum 
LEFT OUTER JOIN EpicorLive10.dbo.JobOpDtl jdt ON jo.JobNum = jdt.JobNum AND jo.OprSeq = jdt.OprSeq 
INNER JOIN EpicorLive10.dbo.Resource rs ON jdt.ResourceID = rs.ResourceID 
LEFT OUTER JOIN	(SELECT ld.JobNum, ld.OprSeq, SUM(ld.LaborQty)[ProdQty]
				 FROM EpicorLive10.dbo.LaborDtl ld
				 WHERE ld.JobNum = ?
				 GROUP BY ld.JobNum, ld.OprSeq) src ON jh.JobNum = src.JobNum AND jdt.OprSeq = src.OprSeq
WHERE jh.JobNum = ? AND jo.OpCode IN ('SWISS','CNC')
