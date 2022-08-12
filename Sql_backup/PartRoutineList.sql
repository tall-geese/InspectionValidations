-- Working Query for selecting only the routines that have Features.
-- Occasionally we have to obsolete a Routine but we can't delete it because it is still relevant to historical data
SELECT rt.RoutineName 
FROM dbo.Part p 
LEFT OUTER JOIN dbo.Feature ft ON p.PartID = ft.PartID 
INNER JOIN dbo.RoutineFeatures rtf ON rtf.FeatureID = ft.FeatureID 
LEFT OUTER JOIN dbo.Routine rt ON rt.RoutineID = rtf.RoutineID 
WHERE p.PartName = ?
GROUP BY rt.RoutineName 