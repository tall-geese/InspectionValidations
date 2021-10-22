

--Grab the beginning TimeStamp and max EmpoyeeID for each observation
--We are also filtering out failures after the grouping
SELECT src2.*
FROM (SELECT MIN(src.ObsTimestamp)[TimeStamp], MAX(src.ItemName)[EmployeeID], src.ObsID, MIN(src.Result)[Result]
	FROM (SELECT frd.ObsTimestamp, dta.ItemName, frd.ObsID, frd.Value,
				CASE 
					WHEN(frd.Value > fpr.UpperToleranceLimit OR frd.Value < fpr.LowerToleranceLimit) THEN 'Fail'
					WHEN frd.Value IS NULL THEN 'Fail'
					ELSE 'Pass'
				END AS 'Result'
			FROM dbo.FeatureRun fr 
			INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
			INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
			INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
			INNER JOIN dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID
			LEFT OUTER JOIN dbo.FeatureProperties fpr ON f.FeatureID = fpr.FeatureID AND f.FeaturePropID = fpr.FeaturePropID 
			LEFT OUTER JOIN dbo.DataTraceability dta ON r.RunID = dta.RunID AND f.FeatureID = dta.FeatureID AND frd.ObsID = dta.StartObsID 
			WHERE r.RunName = ? AND rt.RoutineName = ?
			UNION ALL
			SELECT  afrd.ObsTimestamp, dta.ItemName, afrd.ObsID, afrd.DefectCount,
				CASE
					WHEN afrd.DefectCount = 1 THEN 'Fail'
					WHEN afrd.DefectCount IS NULL THEN 'Fail'
					ELSE 'Pass'
				END AS 'Result'
			FROM dbo.FeatureRun fr 
			INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
			INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
			INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
			INNER JOIN dbo.AttFeatureRunData afrd ON fr.RunID = afrd.RunID AND fr.FeatureID = afrd.FeatureID
			LEFT OUTER JOIN dbo.DataTraceability dta ON r.RunID = dta.RunID AND f.FeatureID = dta.FeatureID AND afrd.ObsID = dta.StartObsID 
			WHERE r.RunName = ? AND rt.RoutineName = ?) src
	GROUP BY src.ObsID) src2
WHERE src2.Result = 'Pass'
ORDER BY src2.ObsID;



--Grabbing ALL Results, including failures
SELECT MIN(src.ObsTimestamp)[TimeStamp], MAX(src.ItemName)[EmployeeID], src.ObsID, MIN(src.Result)[Result]
FROM (SELECT frd.ObsTimestamp, dta.ItemName, frd.ObsID, frd.Value,
			CASE 
				WHEN(frd.Value > fpr.UpperToleranceLimit OR frd.Value < fpr.LowerToleranceLimit) THEN 'Fail'
				WHEN frd.Value IS NULL THEN 'Fail'
				ELSE 'Pass'
			END AS 'Result'
		FROM dbo.FeatureRun fr 
		INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
		INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
		INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
		INNER JOIN dbo.FeatureRunData frd ON fr.RunID = frd.RunID AND fr.FeatureID=frd.FeatureID
		LEFT OUTER JOIN dbo.FeatureProperties fpr ON f.FeatureID = fpr.FeatureID AND f.FeaturePropID = fpr.FeaturePropID 
		LEFT OUTER JOIN dbo.DataTraceability dta ON r.RunID = dta.RunID AND f.FeatureID = dta.FeatureID AND frd.ObsID = dta.StartObsID 
		WHERE r.RunName = ? AND rt.RoutineName = ?
		UNION ALL
		SELECT  afrd.ObsTimestamp, dta.ItemName, afrd.ObsID, afrd.DefectCount,
			CASE
				WHEN afrd.DefectCount = 1 THEN 'Fail'
				WHEN afrd.DefectCount IS NULL THEN 'Fail'
				ELSE 'Pass'
			END AS 'Result'
		FROM dbo.FeatureRun fr 
		INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
		INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
		INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
		INNER JOIN dbo.AttFeatureRunData afrd ON fr.RunID = afrd.RunID AND fr.FeatureID = afrd.FeatureID
		LEFT OUTER JOIN dbo.DataTraceability dta ON r.RunID = dta.RunID AND f.FeatureID = dta.FeatureID AND afrd.ObsID = dta.StartObsID 
		WHERE r.RunName = ? AND rt.RoutineName = ?) src
GROUP BY src.ObsID
ORDER BY src.ObsID