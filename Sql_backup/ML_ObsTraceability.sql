
------------------------------------------
	--- TRACEABILITY DATA -------
-------------------------------------------


--Grab the beginning TimeStamp and max EmpoyeeID for each observation
--We are also filtering out failures after the grouping
SELECT src2.TimeStamp, src2.EmployeeID, src2.ObsNo, 
			CASE 
				WHEN src3.ReqFtCount IS NULL THEN 'Fail'
				ELSE src2.Result
			END AS [Result]
FROM (SELECT MIN(src.ObsTimestamp)[TimeStamp], MAX(src.ItemName)[EmployeeID], src.ObsNo, MIN(src.Result)[Result], MAX(src.RunID)[RunID], COUNT(*)[ObsFtCount]
	FROM (SELECT frd.ObsTimestamp, dta.ItemName, frd.ObsNo, frd.Value, r.RunID,
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
			WHERE r.RunName = ? AND rt.RoutineName = ? AND frd.ObsNo > 0
			UNION ALL
			SELECT  afrd.ObsTimestamp, dta.ItemName, afrd.ObsNo, afrd.DefectCount, r.RunID,
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
			WHERE r.RunName = ? AND rt.RoutineName = ? AND afrd.ObsNo > 0) src
	GROUP BY src.ObsNo) src2
LEFT OUTER JOIN (SELECT r.RunID, COUNT(*)[ReqFtCount]
				FROM .dbo.Run r
				LEFT OUTER JOIN .dbo.Routine rt ON r.RoutineID = rt.RoutineID
				INNER JOIN .dbo.FeatureRun fr ON fr.RunID = r.RunID
				INNER JOIN .dbo.Feature f ON f.FeatureID = fr.FeatureID
				WHERE r.RunName = ? AND rt.RoutineName = ?
				GROUP BY r.RunID) src3 ON src2.RunID = src3.RunID AND src2.ObsFtCount = src3.ReqFtCount 	
WHERE CASE 
		WHEN src3.ReqFtCount IS NULL THEN 'Fail'
		ELSE src2.Result
	END = 'Pass'
ORDER BY src2.ObsNo;


--Grabbing ALL Results, including failures
SELECT src2.TimeStamp, src2.EmployeeID, src2.ObsNo, 
			CASE 
				WHEN src3.ReqFtCount IS NULL THEN 'Fail'
				ELSE src2.Result
			END AS [Result]
FROM (SELECT MIN(src.ObsTimestamp)[TimeStamp], MAX(src.ItemName)[EmployeeID], src.ObsNo, MIN(src.Result)[Result], MAX(src.RunID)[RunID], COUNT(*)[ObsFtCount]
	FROM (SELECT frd.ObsTimestamp, dta.ItemName, frd.ObsNo, frd.Value, r.RunID,
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
			WHERE r.RunName = ? AND rt.RoutineName = ? AND frd.ObsNo > 0
			UNION ALL
			SELECT  afrd.ObsTimestamp, dta.ItemName, afrd.ObsNo, afrd.DefectCount, r.RunID,
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
			WHERE r.RunName = ? AND rt.RoutineName = ? AND afrd.ObsNo > 0) src
	GROUP BY src.ObsNo) src2
LEFT OUTER JOIN (SELECT r.RunID, COUNT(*)[ReqFtCount]
				FROM .dbo.Run r
				LEFT OUTER JOIN .dbo.Routine rt ON r.RoutineID = rt.RoutineID
				INNER JOIN .dbo.FeatureRun fr ON fr.RunID = r.RunID
				INNER JOIN .dbo.Feature f ON f.FeatureID = fr.FeatureID
				WHERE r.RunName = ? AND rt.RoutineName = ?
				GROUP BY r.RunID) src3 ON src2.RunID = src3.RunID AND src2.ObsFtCount = src3.ReqFtCount 	
ORDER BY src2.ObsNo;



--For use in FinalDim, b/c we can't treat observations as fail
--just because they did a single attribute feature. So we are only concerned about the variable features here.
--Also Variable feautres are split off into their own sheet for FI_DIM
SELECT src2.TimeStamp, src2.EmployeeID, src2.ObsNo, 
			CASE 
				WHEN src3.ReqFtCount IS NULL THEN 'Fail'
				ELSE src2.Result
			END AS [Result], src3.*, src2.ObsFtCount
FROM (SELECT MIN(src.ObsTimestamp)[TimeStamp], MAX(src.ItemName)[EmployeeID], src.ObsNo, MIN(src.Result)[Result], MAX(src.RunID)[RunID], COUNT(*)[ObsFtCount]
	FROM (SELECT frd.ObsTimestamp, dta.ItemName, frd.ObsNo, frd.Value, r.RunID,
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
			WHERE r.RunName = ? AND rt.RoutineName = ? AND frd.ObsNo > 0) src
	GROUP BY src.ObsNo) src2
LEFT OUTER JOIN (SELECT r.RunID, COUNT(*)[ReqFtCount]
				FROM .dbo.Run r
				LEFT OUTER JOIN .dbo.Routine rt ON r.RoutineID = rt.RoutineID
				INNER JOIN .dbo.FeatureRun fr ON fr.RunID = r.RunID
				INNER JOIN .dbo.Feature f ON f.FeatureID = fr.FeatureID
				WHERE r.RunName = ? AND rt.RoutineName = ? AND f.FeatureType = 1
				GROUP BY r.RunID) src3 ON src2.RunID = src3.RunID AND src2.ObsFtCount = src3.ReqFtCount 	
WHERE CASE 
		WHEN src3.ReqFtCount IS NULL THEN 'Fail'
		ELSE src2.Result
	END = 'Pass'
ORDER BY src2.ObsNo;
