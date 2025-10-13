--Modify by Eric@20250801
SELECT c.c_comp_co,c.c_y,s.C_SHD_ACTNO,s.C_ACTUISU_DIV,to_char(S.C_WITINTDT,'yyyymmdd'),S.C_WITINT_bono
FROM SHDDRAWEXD s , COMPANYDRAWEXD c
WHERE    s.COMPANYDRAWEXD_UUID = c.uuid
AND c.c_comp_co = :company_code
AND s.c_Witint_Bono <> 0
AND to_char(S.C_WITINTDT,'yyyymmdd') between :start_date and :end_date