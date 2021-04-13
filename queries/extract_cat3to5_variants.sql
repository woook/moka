
-- Query used to extract Category 3 to 5 variants for Congenica transfer: See INC0103359
SELECT dbo.NGSVariant.NGSVariantID, dbo.NGSVariant.DateAdded, dbo.NGSVariant.ChrID, dbo.NGSVariant.Position_hg19, 
dbo.NGSVariant.ref, dbo.NGSVariant.alt, dbo.Status.Status AS 'ACMG_Class',
CASE dbo.NGSVariantACMG.PVS1
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PVS1,
dbo.NGSVariantACMG.PVS1_comment,
CASE dbo.NGSVariantACMG.PS1
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PS1,
dbo.NGSVariantACMG.PS1_comment,
CASE dbo.NGSVariantACMG.PS2
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PS2,
dbo.NGSVariantACMG.PS2_comment,
CASE dbo.NGSVariantACMG.PS3
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PS3,
dbo.NGSVariantACMG.PS3_comment,
CASE dbo.NGSVariantACMG.PS4
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PS4,
dbo.NGSVariantACMG.PS4_comment,
CASE dbo.NGSVariantACMG.PM1
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PM1,
dbo.NGSVariantACMG.PM1_comment,
CASE dbo.NGSVariantACMG.PM2
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PM2,
dbo.NGSVariantACMG.PM2_comment,
CASE dbo.NGSVariantACMG.PM3
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PM3,
dbo.NGSVariantACMG.PM3_comment,
CASE dbo.NGSVariantACMG.PM4
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PM4,
dbo.NGSVariantACMG.PM4_comment,
CASE dbo.NGSVariantACMG.PM5
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PM5,
dbo.NGSVariantACMG.PM5_comment,
CASE dbo.NGSVariantACMG.PM6
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PM6,
dbo.NGSVariantACMG.PM6_comment,
CASE dbo.NGSVariantACMG.PP1
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PP1,
dbo.NGSVariantACMG.PP1_comment,
CASE dbo.NGSVariantACMG.PP2
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PP2,
dbo.NGSVariantACMG.PP2_comment,
CASE dbo.NGSVariantACMG.PP3
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PP3,
dbo.NGSVariantACMG.PP3_comment,
CASE dbo.NGSVariantACMG.PP4
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PP4,
dbo.NGSVariantACMG.PP4_comment,
CASE dbo.NGSVariantACMG.PP5
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS PP5,
dbo.NGSVariantACMG.PP5_comment,
CASE dbo.NGSVariantACMG.BA1
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BA1,
dbo.NGSVariantACMG.BA1_comment,

CASE dbo.NGSVariantACMG.BS1
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BS1,
dbo.NGSVariantACMG.BS1_comment,
CASE dbo.NGSVariantACMG.BS2
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BS2,
dbo.NGSVariantACMG.BS2_comment,
CASE dbo.NGSVariantACMG.BS3
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BS3,
dbo.NGSVariantACMG.BS3_comment,
CASE dbo.NGSVariantACMG.BS4
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BS4,
dbo.NGSVariantACMG.BS4_comment,
CASE dbo.NGSVariantACMG.BP1
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BP1,
dbo.NGSVariantACMG.BP1_comment,
CASE dbo.NGSVariantACMG.BP2
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BP2,
dbo.NGSVariantACMG.BP2_comment,
CASE dbo.NGSVariantACMG.BP3
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BP3,
dbo.NGSVariantACMG.BP3_comment,
CASE dbo.NGSVariantACMG.BP4
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BP4,
dbo.NGSVariantACMG.BP4_comment,
CASE dbo.NGSVariantACMG.BP5
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BP5,
dbo.NGSVariantACMG.BP5_comment,
CASE dbo.NGSVariantACMG.BP6
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BP6,
dbo.NGSVariantACMG.BP6_comment,
CASE dbo.NGSVariantACMG.BP7
	WHEN '3366' THEN 'Very Strong'
	WHEN '3367' THEN 'Strong'
	WHEN '3368' THEN 'Moderate'
	WHEN '3369' THEN 'Supporting'
END AS BP7,
dbo.NGSVariantACMG.BP7_comment
FROM dbo.NGSVariant 
LEFT JOIN dbo.Status 
ON dbo.NGSVariant.Classification = dbo.Status.StatusID 
LEFT JOIN dbo.NGSVariantACMG
ON dbo.NGSVariant.NGSVariantID = dbo.NGSVariantACMG.NGSVariantACMGID
WHERE Classification = 1202218788 OR Classification = 1202218783 OR Classification = 1202218781
