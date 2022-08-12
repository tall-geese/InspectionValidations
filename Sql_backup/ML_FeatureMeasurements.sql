
	--we must generate the list of Features at RunTime, replaceing '{Features}'  with a list in the style of 
	-- [0_001_00],[0_020_00],[0_007_02]
	-- Append to  Where Clause the logic for filtering out failures, but Not NULL values
SELECT Pvt.*
	FROM (SELECT src.FeatureName, src.ObsNo, src.Value
			FROM (SELECT f.FeatureName,  frd.ObsNo, 
			CASE
				WHEN (frd.Value > frp.UpperToleranceLimit) THEN 99.998
		        WHEN (frd.Value < frp.LowerToleranceLimit) THEN 99.998
		        WHEN frd.Value IS NULL THEN 99.998
		        ELSE frd.Value
	    	END AS 'Value'
			FROM dbo.FeatureRun fr 
			INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
			INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
			INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
			INNER JOIN dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID 
			INNER JOIN dbo.FeatureProperties frp ON f.FeatureID = frp.FeatureID AND f.FeaturePropID = frp.FeaturePropID 
			WHERE r.RunName = ? AND rt.RoutineName = ? AND frd.ObsNo > 0
			UNION ALL
			SELECT f.FeatureName, afrd.ObsNo, COALESCE(afrd.DefectCount,1)[Value]
			FROM dbo.FeatureRun fr 
			INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
			INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
			INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
			INNER JOIN dbo.AttFeatureRunData afrd ON fr.RunID = afrd.RunID AND fr.FeatureID = afrd.FeatureID 
			WHERE r.RunName = ? AND rt.RoutineName = ? AND afrd.ObsNo > 0) src) src2
	PIVOT (
		SUM(Value)
		FOR FeatureName IN ({Features})
		)AS Pvt
WHERE ;

--This is our optional query that we will conditionally use when the
	-- 'ShowAllObservations' Toggle button is pressed
SELECT Pvt.*
	FROM (SELECT f.FeatureName,  frd.ObsNo, frd.Value 
		FROM dbo.FeatureRun fr 
		INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
		INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
		INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
		INNER JOIN dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID 
		WHERE r.RunName = ? AND rt.RoutineName = ? AND frd.ObsNo > 0
		UNION ALL
		SELECT f.FeatureName, afrd.ObsNo, afrd.DefectCount 
		FROM dbo.FeatureRun fr 
		INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
		INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
		INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
		INNER JOIN dbo.AttFeatureRunData afrd ON fr.RunID = afrd.RunID AND fr.FeatureID = afrd.FeatureID 
		WHERE r.RunName = ? AND rt.RoutineName = ? AND afrd.ObsNo > 0) src
PIVOT (
	SUM(Value)
	FOR FeatureName IN ({Features})
	)AS Pvt



