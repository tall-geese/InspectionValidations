-----------------------
--Traceability Info--
	--Get the most recent Employee ID and Date that exists for a feature
			--May only return a subset of features if not all have inspections / employeeIDs
	--Again, VBA should check that we have the number of results expected...
			--And that the recordSet is not EOF
SELECT f.FeatureName, dt2.ItemName, afrd.ObsTimestamp
FROM .dbo.AttFeatureRunData afrd
INNER JOIN(SELECT dt.RunID, dt.FeatureID, MAX(dt.StartObsID)[LatestObs]
			FROM .dbo.Run r 
			LEFT OUTER JOIN .dbo.Routine rt ON r.RoutineID = rt.RoutineID 
			INNER JOIN .dbo.DataTraceability dt ON r.RunID = dt.RunID 
			WHERE r.RunName = ? AND rt.RoutineName = ? AND dt.TraceabilityListID = 143
			GROUP BY dt.RunID, dt.FeatureID) src ON src.RunID = afrd.RunID AND src.FeatureID = afrd.FeatureID AND src.LatestObs = afrd.ObsID 
LEFT OUTER JOIN .dbo.DataTraceability dt2 ON src.RunID = dt2.RunID AND src.FeatureID = dt2.FeatureID AND src.LatestObs = dt2.StartObsID 
LEFT OUTER JOIN .dbo.Feature f ON afrd.FeatureID = f.FeatureID 
ORDER BY f.FeatureName ASC