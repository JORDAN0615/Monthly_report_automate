SELECT  c.c_comp_co, c.c_y, s.C_SHD_ACTNO ,
(s.C_EA_BS_CT + s.C_RES_BS_CT + s.C_GENE_WAR_CT + s.C_EMP_WAR_CT + s.C_SPEPER_WAR_CT  + s.C_EMP_BO_SHA)
,to_char(s.C_WITSHA_DT,'yyyymmdd'), s.C_RECPT_SN
FROM shdexrissuancemaster s  , COMPANYISSUANCE c
WHERE s.COMPANYISSUANCE_UUID =  c.UUID
and to_char(s.C_WITSHA_DT,'yyyymmdd') between '20250301' and '20250331'
AND c.c_comp_co = :company_code