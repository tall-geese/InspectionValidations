--P/F Results-- 
	--grabs only the first row of passed inspections.
	--The VBA must check that the Result is 0 and the Count is equal to the number of features that we have from
		--the header info query from above
SELECT SUM(afrd2.DefectCount)[Result], COUNT(*)[Count]
FROM .dbo.AttFeatureRunData afrd2 
INNER JOIN(SELECT r.RunID, rt.RoutineID,  afrd.FeatureID, MAX(afrd.ObsNo)[LatestObs]
		FROM .dbo.Run r 
		LEFT OUTER JOIN .dbo.Routine rt ON r.RoutineID = rt.RoutineID 
		INNER JOIN .dbo.AttFeatureRunData afrd ON r.RunID = afrd.RunID 
		WHERE r.RunName = ? AND rt.RoutineName = ? AND afrd.ObsNo > 0
		GROUP BY r.RunID, rt.RoutineID, afrd.FeatureID) src ON afrd2.RunID = src.RunID AND afrd2.FeatureID = src.FeatureID AND afrd2.ObsNo = src.LatestObs
		