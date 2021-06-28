SELECT rt.RoutineName,
	CASE
		WHEN r.RunStatus = 1 THEN 'New'
		WHEN r.RunStatus = 2 THEN 'Suspended'
		WHEN r.RunStatus = 3 THEN 'Active' 
		WHEN r.RunStatus = 4 THEN 'Closed'
		WHEN r.RunStatus = 12 THEN 'Archived'
		WHEN r.RunStatus = 260 THEN 'Signed'
	END AS [RunStatus]
FROM MeasurLink7.dbo.Run r
LEFT OUTER JOIN MeasurLink7.dbo.Routine rt ON r.RoutineID = rt.RoutineID 
WHERE r.RunName = ?
