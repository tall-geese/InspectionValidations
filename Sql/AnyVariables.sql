SELECT r.RoutineName, f.FeatureType 
FROM MeasurLink7.dbo.Routine r
INNER JOIN MeasurLink7.dbo.RoutineFeatures rf ON r.RoutineID = rf.RoutineID 
INNER JOIN MeasurLink7.dbo.Feature f ON rf.FeatureID = f.FeatureID 
WHERE r.RoutineName = ? AND f.FeatureType = 1