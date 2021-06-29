-- Finished Pivot Table for Feature Measurements, in addition to the RunName and RoutineName Parameters,
	--we must generate the list of Features at RunTime, replaceing '{Features}'  with a list in the style of 
	-- [0_001_00],[0_020_00],[0_007_02]
	-- Since this is in an aggregate, we can't apply logic to the Attribute features, 
		--so when header information is set, we may need to change the result Values to 'Pass/Fail' 
		--depending on the feature type given
SELECT Pvt.*
FROM (SELECT f.FeatureName,  frd.ObsNo, frd.Value 
	FROM MeasurLink7.dbo.FeatureRun fr 
	INNER JOIN MeasurLink7.dbo.Feature f ON F.FeatureID = fr.FeatureID 
	INNER JOIN MeasurLink7.dbo.Run r ON fr.RunID = r.RunID 
	INNER JOIN MeasurLink7.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
	INNER JOIN MeasurLink7.dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID 
	WHERE r.RunName = ? AND rt.RoutineName = ?
	UNION ALL
	SELECT f.FeatureName, afrd.ObsNo, afrd.DefectCount 
	FROM MeasurLink7.dbo.FeatureRun fr 
	INNER JOIN MeasurLink7.dbo.Feature f ON F.FeatureID = fr.FeatureID 
	INNER JOIN MeasurLink7.dbo.Run r ON fr.RunID = r.RunID 
	INNER JOIN MeasurLink7.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
	INNER JOIN MeasurLink7.dbo.AttFeatureRunData afrd ON fr.RunID = afrd.RunID AND fr.FeatureID = afrd.FeatureID 
	WHERE r.RunName = ? AND rt.RoutineName = ?) src	
PIVOT (
	SUM(Value)
	FOR FeatureName IN ({Features})
	)AS Pvt

	
	


-- Attempting a Pivot
-- TODO: finish this up later
--SELECT Pvt.*
--FROM (SELECT f.FeatureName,  frd.ObsNo, frd.Value 
--	FROM MeasurLink7Test.dbo.FeatureRun fr 
--	INNER JOIN MeasurLink7Test.dbo.Feature f ON F.FeatureID = fr.FeatureID 
--	INNER JOIN MeasurLink7Test.dbo.Run r ON fr.RunID = r.RunID 
--	INNER JOIN MeasurLink7Test.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
--	INNER JOIN MeasurLink7Test.dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID 
--	WHERE r.RunName = 'NVXXXX2' AND rt.RoutineName = '1642652_D_IP_EXAMPLE'
--	UNION ALL
--	SELECT f.FeatureName, afrd.ObsNo,
--	, afrd.DefectCount 
--			CASE 
--				WHEN afrd.DefectCount = 1 THEN 'Fail'
--				ELSE 'Pass'
--			END AS [Value]
--	FROM MeasurLink7Test.dbo.FeatureRun fr 
--	INNER JOIN MeasurLink7Test.dbo.Feature f ON F.FeatureID = fr.FeatureID 
--	INNER JOIN MeasurLink7Test.dbo.Run r ON fr.RunID = r.RunID 
--	INNER JOIN MeasurLink7Test.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
--	INNER JOIN MeasurLink7Test.dbo.AttFeatureRunData afrd ON fr.RunID = afrd.RunID AND fr.FeatureID = afrd.FeatureID 
--	WHERE r.RunName = 'NVXXXX2' AND rt.RoutineName = '1642652_D_IP_EXAMPLE') src	
--PIVOT (
--	SUM(Value)
--	FOR FeatureName IN ([0_027_01],[0_026_01],[0_P26_01])
--	)AS Pvt


