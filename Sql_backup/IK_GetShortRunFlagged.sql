

SELECT vru.DrawNum, vru.Revision, DATEDIFF(day, vru.DateApplied, ?)[DaysSinceFlag]
FROM VettingReportsUpdated vru
WHERE vru.DrawNum = ? AND vru.Revision = ?
