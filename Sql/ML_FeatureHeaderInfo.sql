--Get a Distinct list of features applicable to the Run/Routine combination;
--Also grab header information to display on the report including:
--BALLOON (must be formatted), Description, LowerTolerance, Target, UpperTolerance, GaugeID 
--We CAN'T have a float tolerance limit and something like 'N/A' in the same column,
	--gotta switch the logic based on the Type
SELECT f.FeatureName, cfv.ValueString, COALESCE(fpr.LowerToleranceLimit,0)[LowerTolerance], COALESCE(fpr.Target,0)[Target], 
		COALESCE(fpr.UpperToleranceLimit,0)[UpperTolerance], g.GageName, 
		CASE
			WHEN fpr.Target IS NULL THEN 'Attribute'
			ELSE 'Variable'
		END AS [Type]
FROM MeasurLink7.dbo.FeatureRun fr 
INNER JOIN MeasurLink7.dbo.Feature f ON F.FeatureID = fr.FeatureID 
INNER JOIN MeasurLink7.dbo.Run r ON fr.RunID = r.RunID 
INNER JOIN MeasurLink7.dbo.Routine rt ON rt.RoutineID = r.RoutineID 
LEFT OUTER JOIN MeasurLink7.dbo.FeatureProperties fpr ON f.FeatureID = fpr.FeatureID AND f.FeaturePropID = fpr.FeaturePropID 
LEFT OUTER JOIN MeasurLink7.dbo.DataGageTracking dgt ON f.FeatureID = dgt.FeatureID AND r.RunID = dgt.RunID
LEFT OUTER JOIN MeasurLink7.dbo.Gage g ON dgt.GageID = g.GageID 
LEFT OUTER JOIN MeasurLink7.dbo.CustomFieldValue cfv ON f.FeatureID = cfv.ApplyToID
WHERE r.RunName = 'SD1284' AND rt.RoutineName = 'DRW-00717-01_RAJ_IP_IXSHIFT' AND cfv.CustomFieldID = 16