SELECT src.jobNum, MAX(src.AcceptQty)[AcceptQty]
FROM (SELECT ld.JobNum, ld.OprSeq, SUM(ld.LaborQty)[AcceptQty]
		FROM EpicorLive10.dbo.LaborDtl ld 
		WHERE ld.JobNum = ?
		GROUP BY ld.JobNum, ld.OprSeq) src
GROUP BY src.JobNum