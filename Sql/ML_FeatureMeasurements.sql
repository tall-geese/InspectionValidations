
	--we must generate the list of Features at RunTime, replaceing '{Features}'  with a list in the style of 
	-- [0_001_00],[0_020_00],[0_007_02]
	-- Append to  Where Clause the logic for filtering out failures, but Not NULL values
SELECT Pvt.*
	FROM (SELECT src.FeatureName, src.ObsID, src.Value
			FROM (SELECT f.FeatureName,  frd.ObsID, 
			CASE
				WHEN (frd.Value > frp.UpperToleranceLimit) THEN 99.998
		        WHEN (frd.Value < frp.LowerToleranceLimit) THEN 99.998
		        WHEN frd.Value IS NULL THEN 99.998
		        ELSE frd.Value
	    	END AS 'Value'
			FROM MeasurLink7.dbo.FeatureRun fr 
			INNER JOIN MeasurLink7.dbo.Feature f ON F.FeatureID = fr.FeatureID 
			INNER JOIN MeasurLink7.dbo.Run r ON fr.RunID = r.RunID 
			INNER JOIN MeasurLink7.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
			INNER JOIN MeasurLink7.dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID 
			INNER JOIN MeasurLink7.dbo.FeatureProperties frp ON f.FeatureID = frp.FeatureID AND f.FeaturePropID = frp.FeaturePropID 
			WHERE r.RunName = ? AND rt.RoutineName = ?
			UNION ALL
			SELECT f.FeatureName, afrd.ObsID, COALESCE(afrd.DefectCount,1)[Value]
			FROM MeasurLink7.dbo.FeatureRun fr 
			INNER JOIN MeasurLink7.dbo.Feature f ON F.FeatureID = fr.FeatureID 
			INNER JOIN MeasurLink7.dbo.Run r ON fr.RunID = r.RunID 
			INNER JOIN MeasurLink7.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
			INNER JOIN MeasurLink7.dbo.AttFeatureRunData afrd ON fr.RunID = afrd.RunID AND fr.FeatureID = afrd.FeatureID 
			WHERE r.RunName = ? AND rt.RoutineName = ?) src) src2
	PIVOT (
		SUM(Value)
		FOR FeatureName IN ({Features})
		)AS Pvt
WHERE ;

--This is our optional query that we will conditionally use when the
	-- 'ShowAllObservations' Toggle button is pressed
SELECT Pvt.*
	FROM (SELECT f.FeatureName,  frd.ObsID, frd.Value 
		FROM MeasurLink7.dbo.FeatureRun fr 
		INNER JOIN MeasurLink7.dbo.Feature f ON F.FeatureID = fr.FeatureID 
		INNER JOIN MeasurLink7.dbo.Run r ON fr.RunID = r.RunID 
		INNER JOIN MeasurLink7.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
		INNER JOIN MeasurLink7.dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID 
		WHERE r.RunName = ? AND rt.RoutineName = ?
		UNION ALL
		SELECT f.FeatureName, afrd.ObsID, afrd.DefectCount 
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



