

--Test with NV18209

WITH recursiveFolder
	(depth, name)
	AS 
	(SELECT 1, jh.JobNum
	FROM EpicorLive11.dbo.JobHead jh
	WHERE jh.JobNum = ?
	UNION ALL
	SELECT parent.depth + 1, child.JobNum
	FROM recursiveFolder parent 
	INNER JOIN EpicorLive11.dbo.JobHead child ON child.JobNum = ? + '-' + CAST(parent.depth AS NVARCHAR))
SELECT rf.depth, rf.name
FROM recursiveFolder rf