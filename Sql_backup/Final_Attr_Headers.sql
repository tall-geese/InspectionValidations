--Feature and Gauge info--
	--Should return header information even if only gauge information is entered
	--Otherwise the results should be the same, but GageName column will be NULL values
---------------------------
SELECT DISTINCT cfv5.ValueString, cfv.ValueString[Characteristic Desc], cfv2.ValueString[Tool/Tolernace], cfv3.ValueString[Frequency], cfv4.ValueString[Insp Method], f.FeatureName
FROM dbo.FeatureRun fr 
INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
LEFT OUTER JOIN dbo.CustomFieldValue cfv ON f.FeatureID = cfv.ApplyToID
LEFT OUTER JOIN dbo.CustomFieldValue cfv2 ON f.FeatureID = cfv2.ApplyToID
LEFT OUTER JOIN dbo.CustomFieldValue cfv3 ON f.FeatureID = cfv3.ApplyToID
LEFT OUTER JOIN dbo.CustomFieldValue cfv4 ON f.FeatureID = cfv4.ApplyToID
LEFT OUTER JOIN dbo.CustomFieldValue cfv5 ON f.FeatureID = cfv5.ApplyToID
LEFT OUTER JOIN dbo.FeatureProperties fpr ON f.FeatureID = fpr.FeatureID AND f.FeaturePropID = fpr.FeaturePropID 
WHERE r.RunName = ? AND rt.RoutineName = ?
	AND cfv.CustomFieldID = 16 AND cfv2.CustomFieldID = 3  AND fpr.Target IS NULL
	AND cfv3.CustomFieldID = 12 AND cfv4.CustomFieldID = 11 AND cfv5.CustomFieldID = 13
ORDER BY f.FeatureName ASC


-- Feature and Gauge info--
	-- Should return header information even if only gauge information is entered
	-- Otherwise the results should be the same, but GageName column will be NULL values
-------------------------
-- SELECT DISTINCT f.FeatureName, cfv.ValueString[Characteristic Desc], cfv2.ValueString[Tool/Tolernace], cfv3.ValueString[Frequency], g.GageName
-- FROM dbo.FeatureRun fr 
-- INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
-- INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
-- INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
-- LEFT OUTER JOIN dbo.FeatureProperties fpr ON f.FeatureID = fpr.FeatureID AND f.FeaturePropID = fpr.FeaturePropID 
-- LEFT OUTER JOIN dbo.DataGageTracking dgt ON f.FeatureID = dgt.FeatureID AND r.RunID = dgt.RunID
-- LEFT OUTER JOIN dbo.Gage g ON dgt.GageID = g.GageID 
-- LEFT OUTER JOIN dbo.CustomFieldValue cfv ON f.FeatureID = cfv.ApplyToID
-- LEFT OUTER JOIN dbo.CustomFieldValue cfv2 ON f.FeatureID = cfv2.ApplyToID
-- LEFT OUTER JOIN dbo.CustomFieldValue cfv3 ON f.FeatureID = cfv3.ApplyToID
-- WHERE r.RunName = ? AND rt.RoutineName = ?
	-- AND cfv.CustomFieldID = 16 AND cfv2.CustomFieldID = 3 AND COALESCE(dgt.StartObsID, 1) = 1 AND fpr.Target IS NULL
	-- AND cfv3.CustomFieldID = 12
-- ORDER BY f.FeatureName ASC