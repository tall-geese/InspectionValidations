
-- Finished Pivot Table for Feature Measurements, in addition to the RunName and RoutineName Parameters,
	--we must generate the list of Features at RunTime, replaceing '{Features}'  with a list in the style of 
	-- [0_001_00],[0_020_00],[0_007_02]
	-- Since this is in an aggregate, we can't apply logic to the Attribute features, 
		--so when header information is set, we may need to change the result Values to 'Pass/Fail' 
		--depending on the feature type given
SELECT src3.*
FROM (SELECT Pvt.*
	FROM (SELECT src.FeatureName, src.ObsNo, src.Value
			FROM (SELECT f.FeatureName,  frd.ObsNo, frd.Value,
				CASE 
					WHEN (frd.Value > frp.UpperToleranceLimit) THEN 'Fail'
					WHEN (frd.Value < frp.LowerToleranceLimit) THEN 'Fail'
					ELSE 'Pass'
				END AS 'Result'
			FROM MeasurLink7.dbo.FeatureRun fr 
			INNER JOIN MeasurLink7.dbo.Feature f ON F.FeatureID = fr.FeatureID 
			INNER JOIN MeasurLink7.dbo.Run r ON fr.RunID = r.RunID 
			INNER JOIN MeasurLink7.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
			INNER JOIN MeasurLink7.dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID 
			INNER JOIN MeasurLink7.dbo.FeatureProperties frp ON f.FeatureID = frp.FeatureID AND f.FeaturePropID = frp.FeaturePropID 
			WHERE r.RunName = ? AND rt.RoutineName = ?
			UNION ALL
			SELECT f.FeatureName, afrd.ObsNo, afrd.DefectCount,
				CASE 
					WHEN (afrd.DefectCount = 1) THEN 'Fail'
					ELSE 'Pass'
				END AS 'Result'
			FROM MeasurLink7.dbo.FeatureRun fr 
			INNER JOIN MeasurLink7.dbo.Feature f ON F.FeatureID = fr.FeatureID 
			INNER JOIN MeasurLink7.dbo.Run r ON fr.RunID = r.RunID 
			INNER JOIN MeasurLink7.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
			INNER JOIN MeasurLink7.dbo.AttFeatureRunData afrd ON fr.RunID = afrd.RunID AND fr.FeatureID = afrd.FeatureID 
			WHERE r.RunName = ? AND rt.RoutineName = ?) src
			WHERE src.Result = 'Pass') src2
	PIVOT (
		SUM(Value)
		FOR FeatureName IN ({Features})
		)AS Pvt) src3
WHERE ;

--This is our optional query that we will conditionally use when the
	-- 'ShowAllObservations' Toggle button is pressed
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



