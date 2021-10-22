SELECT r.RoutineName, f.FeatureType 
FROM dbo.Routine r
INNER JOIN dbo.RoutineFeatures rf ON r.RoutineID = rf.RoutineID 
INNER JOIN dbo.Feature f ON rf.FeatureID = f.FeatureID 
WHERE r.RoutineName = ? AND f.FeatureType = 1