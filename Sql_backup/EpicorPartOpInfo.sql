-- Find out how many, if any, machining ops the part is supposed to have
SELECT jh.JobNum, po.OprSeq , po.OpCode 
FROM EpicorLive10.dbo.JobHead jh 
LEFT OUTER JOIN EpicorLive10.dbo.PartOpr po ON jh.PartNum = po.PartNum AND jh.RevisionNum = po.RevisionNum 
WHERE jh.JobNum = ? AND po.OpCode IN ('CNC','SWISS')
ORDER BY po.OprSeq ASC