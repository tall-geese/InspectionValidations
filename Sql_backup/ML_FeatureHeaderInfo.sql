--Get a Distinct list of features applicable to the Run/Routine combination;
--Also grab header information to display on the report including:
--BALLOON (must be formatted), Description, LowerTolerance, Target, UpperTolerance, GaugeID 
--We CAN'T have a float tolerance limit and something like 'N/A' in the same column,
	--gotta switch the logic based on the Type
	
SELECT DISTINCT f.FeatureName, cfv.ValueString, COALESCE(fpr.LowerToleranceLimit,0)[LowerTolerance], COALESCE(fpr.Target,0)[Target], 
		COALESCE(fpr.UpperToleranceLimit,0)[UpperTolerance], cfv2.ValueString, 
		CASE
			WHEN fpr.Target IS NULL THEN 'Attribute'
			ELSE 'Variable'
		END AS [Type],
		CASE 
			WHEN CHARINDEX('MAX', cfv3.ValueString) <> 0 THEN CONCAT('/', cfv3.ValueString)
			WHEN CHARINDEX('MIN', cfv3.ValueString) <> 0 THEN CONCAT(cfv3.ValueString, '/')
			WHEN CHARINDEX('/', cfv3.ValueString) = 0 THEN 'NA'
			ELSE cfv3.ValueString
		END AS [FixedAttrTol], cfv4.ValueString[Balloon Num]
FROM dbo.FeatureRun fr 
INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
LEFT OUTER JOIN dbo.FeatureProperties fpr ON fr.FeatureID = fpr.FeatureID AND fr.FeaturePropID = fpr.FeaturePropID 
LEFT OUTER JOIN dbo.CustomFieldValue cfv ON f.FeatureID = cfv.ApplyToID
LEFT OUTER JOIN dbo.CustomFieldValue cfv2 ON f.FeatureID = cfv2.ApplyToID
LEFT OUTER JOIN dbo.CustomFieldValue cfv3 ON f.FeatureID = cfv3.ApplyToID
LEFT OUTER JOIN dbo.CustomFieldValue cfv4 ON f.FeatureID = cfv4.ApplyToID
WHERE r.RunName = ? AND rt.RoutineName = ? AND cfv.CustomFieldID = 16 AND cfv2.CustomFieldID = 11 AND cfv3.CustomFieldID = 3 AND cfv4.CustomFieldID = 13 AND f.FeatureName NOT LIKE '%_DEV' 
	
	

	
	
	
-- SELECT DISTINCT f.FeatureName, cfv.ValueString, COALESCE(fpr.LowerToleranceLimit,0)[LowerTolerance], COALESCE(fpr.Target,0)[Target], 
		-- COALESCE(fpr.UpperToleranceLimit,0)[UpperTolerance], g.GageName, 
		-- CASE
			-- WHEN fpr.Target IS NULL THEN 'Attribute'
			-- ELSE 'Variable'
		-- END AS [Type]
-- FROM dbo.FeatureRun fr 
-- INNER JOIN dbo.Feature f ON F.FeatureID = fr.FeatureID 
-- INNER JOIN dbo.Run r ON fr.RunID = r.RunID 
-- INNER JOIN dbo.Routine rt ON rt.RoutineID = r.RoutineID 
-- LEFT OUTER JOIN dbo.FeatureProperties fpr ON f.FeatureID = fpr.FeatureID AND f.FeaturePropID = fpr.FeaturePropID 
-- LEFT OUTER JOIN dbo.DataGageTracking dgt ON f.FeatureID = dgt.FeatureID AND r.RunID = dgt.RunID
-- LEFT OUTER JOIN dbo.Gage g ON dgt.GageID = g.GageID 
-- LEFT OUTER JOIN dbo.CustomFieldValue cfv ON f.FeatureID = cfv.ApplyToID
-- WHERE r.RunName = ? AND rt.RoutineName = ? AND cfv.CustomFieldID = 16 AND dgt.StartObsID = 1