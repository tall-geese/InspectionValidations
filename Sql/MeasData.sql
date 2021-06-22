-- Working Query, Attr and Variable Features are tracked in different Tables
-- So we need to UNION the results
SELECT r.RunName, f.FeatureName, fr.FeatureID, rt.RoutineName, frd.ObsNo, frd.Value, frd.ObsFlags 
FROM MeasurLink7.dbo.FeatureRun fr 
INNER JOIN MeasurLink7.dbo.Feature f ON F.FeatureID = fr.FeatureID 
INNER JOIN MeasurLink7.dbo.Run r ON fr.RunID = r.RunID 
INNER JOIN MeasurLink7.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
INNER JOIN MeasurLink7.dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID 
WHERE r.RunName = 'SD1284' AND rt.RoutineName = 'DRW-00717-01_RAJ_IP_IXSHIFT'
UNION ALL
SELECT r.RunName, f.FeatureName, fr.FeatureID, rt.RoutineName, afrd.ObsNo, afrd.DefectCount, afrd.ObsFlags 
FROM MeasurLink7.dbo.FeatureRun fr 
INNER JOIN MeasurLink7.dbo.Feature f ON F.FeatureID = fr.FeatureID 
INNER JOIN MeasurLink7.dbo.Run r ON fr.RunID = r.RunID 
INNER JOIN MeasurLink7.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
INNER JOIN MeasurLink7.dbo.AttFeatureRunData afrd ON fr.RunID = afrd.RunID AND fr.FeatureID = afrd.FeatureID 
WHERE r.RunName = 'SD1284' AND rt.RoutineName = 'DRW-00717-01_RAJ_IP_IXSHIFT' 
ORDER BY ObsNo, FeatureName

-- USING THE ML TEST DATABASE
SELECT r.RunName, f.FeatureName, fr.FeatureID, rt.RoutineName, frd.ObsNo, frd.Value, frd.ObsFlags 
FROM MeasurLink7Test.dbo.FeatureRun fr 
INNER JOIN MeasurLink7Test.dbo.Feature f ON F.FeatureID = fr.FeatureID 
INNER JOIN MeasurLink7Test.dbo.Run r ON fr.RunID = r.RunID 
INNER JOIN MeasurLink7Test.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
INNER JOIN MeasurLink7Test.dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID 
WHERE r.RunName = 'NVXXXX2' AND rt.RoutineName = '1642652_D_IP_EXAMPLE'
UNION ALL
SELECT r.RunName, f.FeatureName, fr.FeatureID, rt.RoutineName, afrd.ObsNo, afrd.DefectCount, afrd.ObsFlags 
FROM MeasurLink7Test.dbo.FeatureRun fr 
INNER JOIN MeasurLink7Test.dbo.Feature f ON F.FeatureID = fr.FeatureID 
INNER JOIN MeasurLink7Test.dbo.Run r ON fr.RunID = r.RunID 
INNER JOIN MeasurLink7Test.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
INNER JOIN MeasurLink7Test.dbo.AttFeatureRunData afrd ON fr.RunID = afrd.RunID AND fr.FeatureID = afrd.FeatureID 
WHERE r.RunName = 'NVXXXX2' AND rt.RoutineName = '1642652_D_IP_EXAMPLE' 

----------------------------------------------------------------

-- Attempting a Pivot
-- TODO: finish this up later
SELECT src.FeatureName, src.ObsNo
FROM (SELECT r.RunName, f.FeatureName, fr.FeatureID, rt.RoutineName, frd.ObsNo, frd.Value, frd.ObsFlags 
FROM MeasurLink7Test.dbo.FeatureRun fr 
INNER JOIN MeasurLink7Test.dbo.Feature f ON F.FeatureID = fr.FeatureID 
INNER JOIN MeasurLink7Test.dbo.Run r ON fr.RunID = r.RunID 
INNER JOIN MeasurLink7Test.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
INNER JOIN MeasurLink7Test.dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID 
WHERE r.RunName = 'NVXXXX2' AND rt.RoutineName = '1642652_D_IP_EXAMPLE'
UNION ALL
SELECT r.RunName, f.FeatureName, fr.FeatureID, rt.RoutineName, afrd.ObsNo, afrd.DefectCount, afrd.ObsFlags 
FROM MeasurLink7Test.dbo.FeatureRun fr 
INNER JOIN MeasurLink7Test.dbo.Feature f ON F.FeatureID = fr.FeatureID 
INNER JOIN MeasurLink7Test.dbo.Run r ON fr.RunID = r.RunID 
INNER JOIN MeasurLink7Test.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
INNER JOIN MeasurLink7Test.dbo.AttFeatureRunData afrd ON fr.RunID = afrd.RunID AND fr.FeatureID = afrd.FeatureID 
WHERE r.RunName = 'NVXXXX2' AND rt.RoutineName = '1642652_D_IP_EXAMPLE') src
PIVOT AVG(src.Value)
		FOR src.ObsNo AS Pivot_Table

